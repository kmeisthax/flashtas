use com::runtime::init_runtime;
use convert_case::{Case, Casing};
use std::env::args;
use std::ffi::OsStr;
use windows::core::{Error as WinError, HRESULT, HSTRING};
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::{
    ITypeInfo, ITypeLib, FUNCDESC, TKIND_COCLASS, TKIND_DISPATCH, TKIND_INTERFACE, TLIBATTR,
    TYPEATTR,
};
use windows::Win32::System::Ole::LoadTypeLib;

/// Export a type library to Rust source code.
fn print_type_lib_as_rust(lib: &ITypeLib) -> Result<(), WinError> {
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

    println!("        //TODO: {}", strname);

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

            println!("    #[uuid(\"{:?}\")]", typeattr.guid);
            //TODO: Get superinterfaces if possible
            println!("    pub unsafe interface {}: IUnknown {{", strname);

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

        print_type_lib_as_rust(&fp_lib).unwrap();

        println!("use com::interfaces::IUnknown;");
        println!("use com::sys::GUID;");
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
