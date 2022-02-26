//! Exporters for single COM functions

use crate::type_bridge;
use windows::core::{Error as WinError, HRESULT, HSTRING};
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::{
    ITypeInfo, ELEMDESC, FUNCDESC, FUNC_DISPATCH, FUNC_PUREVIRTUAL, FUNC_VIRTUAL, INVOKE_FUNC,
};
use windows::Win32::System::Ole::{VARENUM, VT_VOID};

/// Print valid Rust source code that matches the type signature of a COM
/// method.
fn rust_fn_for_com_method(
    type_nfo: &ITypeInfo,
    funcdesc: &FUNCDESC,
    fn_name: BSTR,
    has_impl: bool,
) -> Result<String, WinError> {
    let mut param_types = vec!["&self".to_string()];
    for i in 0..funcdesc.cParams {
        let elemdesc: &mut ELEMDESC =
            unsafe { &mut *funcdesc.lprgelemdescParam.offset(i as isize) };
        param_types.push(format!(
            "param{}: {}",
            i,
            type_bridge::bridge_elem_to_rust_type(type_nfo, &elemdesc.tdesc)?
        ));
    }

    let return_type = if VARENUM(funcdesc.elemdescFunc.tdesc.vt as i32) == VT_VOID {
        "".to_string()
    } else {
        format!(
            " -> {}",
            type_bridge::bridge_elem_to_rust_type(type_nfo, &funcdesc.elemdescFunc.tdesc)?
        )
    };

    Ok(format!(
        "pub unsafe fn {}({}){}{}",
        fn_name,
        param_types.join(", "),
        return_type,
        if has_impl { "" } else { ";" }
    ))
}

/// Export a single function from a type info structure to Rust source code.
///
/// This is intended to be called within an `interface` block.
pub fn print_type_function_as_rust(type_nfo: &ITypeInfo, fn_index: u32) -> Result<(), WinError> {
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
                rust_fn_for_com_method(type_nfo, funcdesc, strname, false)?
            );
        }
        FUNC_VIRTUAL | FUNC_PUREVIRTUAL => {
            println!(
                "        //INVALID invkind: {} (funckind {:?}, invkind {:?})",
                strname, funcdesc.funckind, funcdesc.invkind
            );
        }
        FUNC_DISPATCH => {}
        _ => {
            println!(
                "        //UNKNOWN funckind: {} (funckind {:?}, invkind {:?})",
                strname, funcdesc.funckind, funcdesc.invkind
            );
        }
    }

    unsafe { type_nfo.ReleaseFuncDesc(funcdesc) };

    Ok(())
}

/// Export a single dispatch property from a type info structure to Rust source
/// code.
///
/// This is intended to be called within a Rust `impl` block.
pub fn print_type_dispatch_as_rust(type_nfo: &ITypeInfo, fn_index: u32) -> Result<(), WinError> {
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
        FUNC_DISPATCH if funcdesc.invkind == INVOKE_FUNC => {
            println!(
                "    {} {{",
                rust_fn_for_com_method(type_nfo, funcdesc, strname, true)?
            );

            println!("        let mut arg_params = vec![];");

            for i in 0..funcdesc.cParams {
                let elemdesc: &mut ELEMDESC =
                    unsafe { &mut *funcdesc.lprgelemdescParam.offset(i as isize) };

                println!(
                    "        {}",
                    type_bridge::generate_dispatch_param(i, &elemdesc.tdesc)?
                );
            }

            println!("        let mut disp_params = DISPPARAMS {{");
            println!("            rgvarg: arg_params.as_mut_ptr(),");
            println!("            rgdispidNamedArgs: ::std::ptr::null_mut(),");
            println!("            cArgs: 0,");
            println!("            cNamedArgs: 0");
            println!("        }};");

            if VARENUM(funcdesc.elemdescFunc.tdesc.vt as i32) != VT_VOID {
                println!("        let mut disp_result = VARIANT::default();");
                println!(
                    "        IDispatch::Invoke(
            &self,
            0x{:X},
            &mut GUID {{
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8]
            }},
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut()
        );",
                    funcdesc.memid
                );

                println!(
                    "        {}",
                    type_bridge::generate_dispatch_return(&funcdesc.elemdescFunc.tdesc)?
                );
            } else {
                println!(
                    "        IDispatch::Invoke(
            &self,
            0x{:X},
            &mut GUID {{
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8]
            }},
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut()
        );",
                    funcdesc.memid
                );
            }

            println!("    }}");
            println!();
        }
        FUNC_DISPATCH => {
            println!(
                "        //TODO: IDispatch helper for {} (invkind {:?})",
                strname, funcdesc.invkind
            );
        }
        _ => {}
    }

    unsafe { type_nfo.ReleaseFuncDesc(funcdesc) };

    Ok(())
}