use com::runtime::init_runtime;
use convert_case::{Case, Casing};
use std::env::args;
use std::ffi::OsStr;
use std::ptr::null_mut;
use windows::core::{Error as WinError, HRESULT, HSTRING};
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::{
    ITypeInfo, ITypeLib, ELEMDESC, FUNCDESC, FUNC_DISPATCH, FUNC_PUREVIRTUAL, FUNC_VIRTUAL,
    INVOKE_FUNC, TKIND_COCLASS, TKIND_DISPATCH, TKIND_INTERFACE, TLIBATTR, TYPEATTR, TYPEDESC,
};
use windows::Win32::System::Ole::LoadTypeLib;
use windows::Win32::System::Ole::{
    VARENUM, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_HRESULT, VT_I1, VT_I2, VT_I4, VT_I8, VT_INT, VT_PTR,
    VT_R4, VT_R8, VT_UI1, VT_UI2, VT_UI4, VT_UI8, VT_UINT, VT_UNKNOWN, VT_USERDEFINED, VT_VARIANT,
    VT_VOID,
};

/// Export a type library's metadata as a module doccomment.
fn print_type_lib_doccomment(lib: &ITypeLib) -> Result<(), WinError> {
    let libattr_raw = unsafe { lib.GetLibAttr()? };
    if libattr_raw.is_null() {
        return Err(WinError::new(HRESULT(-1), HSTRING::new()));
    }

    let libattr: &mut TLIBATTR = unsafe { &mut *libattr_raw };

    println!("//! Exported from COM metadata to Rust via comexport-rs");
    println!("//!");
    println!("//! GUID: {:?}", libattr.guid);
    println!(
        "//! Version: {}.{}",
        libattr.wMajorVerNum, libattr.wMinorVerNum
    );
    println!("//! Locale ID: {}", libattr.lcid);
    println!("//! SYSKIND: {}", libattr.syskind.0);
    println!("//! Flags: {}", libattr.wLibFlags);
    println!();

    //TODO: I think this is unsound
    unsafe { lib.ReleaseTLibAttr(libattr_raw) };

    Ok(())
}

/// Given a type and a referred type ID, print the type it would be if defined
/// in Rust.
///
/// This is sometimes referred to as `HREFTYPE` in Microsoft documentation.
///
/// NOTE: User-defined types are used verbatim. There is currently no machinery
/// to look up or add use statements for user-defined types.
fn bridge_usertype_to_rust_type(typeinfo: &ITypeInfo, href: u32) -> Result<String, WinError> {
    let mut strname = BSTR::new();

    unsafe {
        let target_type = typeinfo.GetRefTypeInfo(href);
        if target_type.is_err() {
            return Ok(format!(
                "/* unknown user-defined type 0x{:X} (error on reftypeinfo) */",
                href
            ));
        }

        let target_type = target_type.unwrap();
        let mut target_lib = None;
        let mut target_index = 0;
        if target_type
            .GetContainingTypeLib(&mut target_lib, &mut target_index)
            .is_err()
        {
            return Ok(format!(
                "/* unknown user-defined type 0x{:X} (error on getcontainingtypelib) */",
                href
            ));
        }

        let target_lib = target_lib.unwrap();
        if target_lib
            .GetDocumentation(
                target_index as i32,
                &mut strname,
                null_mut(),
                null_mut(),
                null_mut(),
            )
            .is_err()
        {
            return Ok(format!(
                "/* unknown user-defined type 0x{:X} (error on getdocs) */",
                href
            ));
        }
    }

    Ok(format!("{}", strname))
}

/// Given a type description and the type it came from, print the type it would
/// be if defined in Rust.
///
/// NOTE: User-defined types are used verbatim. There is currently no machinery
/// to look up or add use statements for user-defined types.
fn bridge_elem_to_rust_type(typeinfo: &ITypeInfo, tdesc: &TYPEDESC) -> Result<String, WinError> {
    Ok(match VARENUM(tdesc.vt as i32) {
        VT_I2 => "i16".to_string(),
        VT_I4 => "i32".to_string(),
        VT_R4 => "f32".to_string(),
        VT_R8 => "f64".to_string(),
        VT_BSTR => "BSTR".to_string(),
        VT_DISPATCH => "IDispatch".to_string(),
        VT_BOOL => "BOOL".to_string(),
        VT_I1 => "i8".to_string(),
        VT_UI1 => "u8".to_string(),
        VT_UI2 => "u16".to_string(),
        VT_UI4 => "u32".to_string(),
        VT_I8 => "i64".to_string(),
        VT_UI8 => "u64".to_string(),
        VT_INT => "isize".to_string(),
        VT_UINT => "usize".to_string(),
        VT_VOID => "c_void".to_string(),
        VT_HRESULT => "HRESULT".to_string(),
        VT_UNKNOWN => "IUnknown".to_string(),
        VT_PTR => {
            let target_type: &TYPEDESC = unsafe { &*tdesc.Anonymous.lptdesc };
            format!("*mut {}", bridge_elem_to_rust_type(typeinfo, target_type)?)
        }
        VT_USERDEFINED | VT_VARIANT => {
            let href_type = unsafe { tdesc.Anonymous.hreftype };
            bridge_usertype_to_rust_type(typeinfo, href_type)?
        }
        _ => format!("/* unknown type 0x{:X} */", tdesc.vt),
    })
}

/// Print valid Rust source code that matches the type signature of a COM
/// method.
fn rust_fn_for_com_method(
    type_nfo: &ITypeInfo,
    funcdesc: &FUNCDESC,
    fn_name: BSTR,
) -> Result<String, WinError> {
    let mut param_types = vec!["&self".to_string()];
    for i in 0..funcdesc.cParams {
        let elemdesc: &mut ELEMDESC =
            unsafe { &mut *funcdesc.lprgelemdescParam.offset(i as isize) };
        param_types.push(format!(
            "param{}: {}",
            i,
            bridge_elem_to_rust_type(type_nfo, &elemdesc.tdesc)?
        ));
    }

    let return_type = if VARENUM(funcdesc.elemdescFunc.tdesc.vt as i32) == VT_VOID {
        "".to_string()
    } else {
        format!(
            " -> {}",
            bridge_elem_to_rust_type(type_nfo, &funcdesc.elemdescFunc.tdesc)?
        )
    };

    Ok(format!(
        "fn {}({}){};",
        fn_name,
        param_types.join(", "),
        return_type
    ))
}

/// Export a single function from a type info structure to Rust source code.
fn print_type_function_as_rust(type_nfo: &ITypeInfo, fn_index: u32) -> Result<(), WinError> {
    let funcdesc_raw = unsafe { type_nfo.GetFuncDesc(fn_index)? };
    if funcdesc_raw.is_null() {
        return Err(WinError::new(HRESULT(-1), HSTRING::new()));
    }

    let funcdesc: &mut FUNCDESC = unsafe { &mut *funcdesc_raw };

    let mut strname = BSTR::new();
    let mut strdocstring = BSTR::new();
    let mut whelpcontext = 0;
    let mut strhelpfile = BSTR::new();

    //TODO: This leaks.
    unsafe {
        type_nfo.GetDocumentation(
            funcdesc.memid,
            &mut strname,
            &mut strdocstring,
            &mut whelpcontext,
            &mut strhelpfile,
        )?
    };

    //TODO: CALLCONV?
    match funcdesc.funckind {
        FUNC_VIRTUAL | FUNC_PUREVIRTUAL if funcdesc.invkind == INVOKE_FUNC => {
            println!(
                "        {}",
                rust_fn_for_com_method(type_nfo, funcdesc, strname)?
            );
        }
        //TODO: Dispatch methods don't live in the interface vtable and should
        //be implemented in another block
        FUNC_DISPATCH if funcdesc.invkind == INVOKE_FUNC => {
            println!("        // TODO: dispatch");
            println!(
                "        // {}",
                rust_fn_for_com_method(type_nfo, funcdesc, strname)?
            );
            println!();
        }
        _ => {
            println!(
                "        //TODO: {} (funckind {:?}, invkind {:?})",
                strname, funcdesc.funckind, funcdesc.invkind
            );
        }
    }

    unsafe { type_nfo.ReleaseFuncDesc(funcdesc) };

    Ok(())
}

/// Export a single type from a type library to Rust source code.
///
/// Since the COM package requires a separate block for interfaces, the
/// `in_interface_block` variable signals if we are printing the interface
/// block or not. Constants should not be printed inside the interface block.
fn print_type_lib_type_as_rust(
    lib: &ITypeLib,
    type_index: u32,
    in_interface_block: bool,
) -> Result<(), WinError> {
    let type_nfo = unsafe { lib.GetTypeInfo(type_index)? };
    let typeattr_raw = unsafe { type_nfo.GetTypeAttr()? };
    if typeattr_raw.is_null() {
        return Err(WinError::new(HRESULT(-1), HSTRING::new()));
    }

    let typeattr: &mut TYPEATTR = unsafe { &mut *typeattr_raw };

    let mut strname = BSTR::new();
    let mut strdocstring = BSTR::new();
    let mut whelpcontext = 0;
    let mut strhelpfile = BSTR::new();

    //TODO: This leaks.
    unsafe {
        lib.GetDocumentation(
            type_index as i32,
            &mut strname,
            &mut strdocstring,
            &mut whelpcontext,
            &mut strhelpfile,
        )?
    };

    match typeattr.typekind {
        //TODO: Other types of COM types
        TKIND_INTERFACE | TKIND_DISPATCH if in_interface_block => {
            if !strdocstring.is_empty() {
                println!("    /// {}", strdocstring);
            }

            let mut superinterfaces = vec![];
            for i in 0..typeattr.cImplTypes {
                let href = unsafe { type_nfo.GetRefTypeOfImplType(i as u32)? };
                superinterfaces.push(bridge_usertype_to_rust_type(&type_nfo, href)?);
            }

            println!("    #[uuid(\"{:?}\")]", typeattr.guid);
            println!(
                "    pub unsafe interface {}: {} {{",
                strname,
                superinterfaces.join(", ")
            );

            for i in 0..typeattr.cFuncs {
                print_type_function_as_rust(&type_nfo, i as u32)?;
            }

            println!("    }}");
            println!();
        }
        TKIND_COCLASS if !in_interface_block => {
            if !strdocstring.is_empty() {
                println!("/// {}", strdocstring);
                println!("///");
            }

            println!("/// GUID: {:?}", typeattr.guid);

            println!(
                "pub const {}_CLSID: GUID = GUID {{",
                strname.to_string().to_case(Case::UpperSnake)
            );
            println!("    data1: 0x{:X},", typeattr.guid.data1);
            println!("    data2: 0x{:X},", typeattr.guid.data2);
            println!("    data3: 0x{:X},", typeattr.guid.data3);
            println!(
                "    data4: [0x{:X}, 0x{:X}, 0x{:X}, 0x{:X}, 0x{:X}, 0x{:X}, 0x{:X}, 0x{:X}],",
                typeattr.guid.data4[0],
                typeattr.guid.data4[1],
                typeattr.guid.data4[2],
                typeattr.guid.data4[3],
                typeattr.guid.data4[4],
                typeattr.guid.data4[5],
                typeattr.guid.data4[6],
                typeattr.guid.data4[7]
            );
            println!("}};");
            println!();
        }
        TKIND_INTERFACE | TKIND_DISPATCH if !in_interface_block => {}
        TKIND_COCLASS if in_interface_block => {}
        k if !in_interface_block => {
            println!("//WARN: Unknown type {} of kind {:?}", strname, k);
        }
        _ => {}
    }

    unsafe { type_nfo.ReleaseTypeAttr(typeattr_raw) };

    Ok(())
}

fn main() {
    init_runtime().expect("A working COM runtime");

    let path = args().nth(1).expect("Path to .exe, .dll, or .ocx");
    let fp_ocx_name = OsStr::new(&path);
    let fp_lib = unsafe { LoadTypeLib(fp_ocx_name).expect("Loaded type lib") };

    unsafe {
        eprintln!("{} has {} entries", path, fp_lib.GetTypeInfoCount());

        print_type_lib_doccomment(&fp_lib).unwrap();

        println!("#![allow(clippy::too_many_arguments)]");
        println!("#![allow(clippy::upper_case_acronyms)]");
        println!();
        println!("use com::interfaces::IUnknown;");

        //TODO: automatic bridging from user-defined types to `windows`/`com` types
        println!("use com::sys::GUID;");
        println!("use windows::core::HRESULT;");
        println!("use windows::Win32::System::Com::{{DISPPARAMS, EXCEPINFO}};");
        println!();
        println!("type BSTR = *const u16;");
        println!();

        for i in 0..fp_lib.GetTypeInfoCount() {
            print_type_lib_type_as_rust(&fp_lib, i, false).unwrap();
        }

        println!("com::interfaces! {{");

        for i in 0..fp_lib.GetTypeInfoCount() {
            print_type_lib_type_as_rust(&fp_lib, i, true).unwrap();
        }

        println!("}}");

        eprintln!("Type definition export complete!");
    }
}
