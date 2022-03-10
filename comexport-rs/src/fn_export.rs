//! Exporters for single COM functions

use crate::context::Context;
use crate::error::Error;
use crate::{dispatch_bridge, type_bridge};
use std::fmt::Write;
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::{
    ITypeInfo, ELEMDESC, FUNCDESC, FUNC_DISPATCH, FUNC_PUREVIRTUAL, FUNC_VIRTUAL, INVOKE_FUNC,
    INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF,
};
use windows::Win32::System::Ole::{VARENUM, VT_VOID};

/// Print valid Rust source code that matches the type signature of a COM
/// method.
///
/// COM methods are meant to correspond directly to the vtable that coclasses
/// expose. In a type library this includes `FUNC_VIRTUAL` and
/// `FUNC_PUREVIRTUAL`. They go into the `interface` block of a given
/// interface.
fn rust_fn_for_com_method(
    context: &mut Context<'_>,
    type_nfo: &ITypeInfo,
    funcdesc: &FUNCDESC,
    fn_name: BSTR,
) -> Result<String, Error> {
    let mut param_types = vec!["&self".to_string()];
    for i in 0..funcdesc.cParams {
        let elemdesc: &mut ELEMDESC =
            unsafe { &mut *funcdesc.lprgelemdescParam.offset(i as isize) };
        param_types.push(format!(
            "param{}: {}",
            i,
            type_bridge::bridge_elem_to_rust_type(context, type_nfo, &elemdesc.tdesc)
        ));
    }

    let return_type = if VARENUM(funcdesc.elemdescFunc.tdesc.vt as i32) == VT_VOID {
        "".to_string()
    } else {
        format!(
            " -> {}",
            type_bridge::bridge_elem_to_rust_type(context, type_nfo, &funcdesc.elemdescFunc.tdesc)
        )
    };

    Ok(format!(
        "pub unsafe fn {}({}){};",
        fn_name,
        param_types.join(", "),
        return_type
    ))
}

/// Export a single function from a type info structure to Rust source code.
pub fn type_function_as_rust(
    context: &mut Context<'_>,
    type_nfo: &ITypeInfo,
    fn_index: u32,
) -> Result<String, Error> {
    let mut ret = String::new();

    let funcdesc_raw = unsafe { type_nfo.GetFuncDesc(fn_index)? };
    if funcdesc_raw.is_null() {
        return Err(Error::NoFuncDescForComFn);
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
            write!(
                ret,
                "        {}",
                rust_fn_for_com_method(context, type_nfo, funcdesc, strname)?
            )?;
        }
        FUNC_VIRTUAL | FUNC_PUREVIRTUAL => {
            write!(
                ret,
                "        //INVALID invkind: {} (funckind {:?}, invkind {:?})",
                strname, funcdesc.funckind, funcdesc.invkind
            )?;
        }
        FUNC_DISPATCH => {}
        _ => {
            write!(
                ret,
                "        //UNKNOWN funckind: {} (funckind {:?}, invkind {:?})",
                strname, funcdesc.funckind, funcdesc.invkind
            )?;
        }
    }

    unsafe { type_nfo.ReleaseFuncDesc(funcdesc) };

    Ok(ret)
}

/// Print valid Rust source code that matches the type signature of a COM
/// dispatch helper.
///
/// COM dispatch helpers exist to allow Rust code to avoid having to handle
/// parameter marshalling for `IDispatch`-based methods and properties. They
/// live in a separate `impl` block for an already-defined COM interface.
///
/// All dispatch helpers return a `Result` whose error type is `HRESULT`. The
/// error fork excludes `S_OK` as a value; that is treated as the success fork
/// of the `Result` which returns the bridged Rust type the caller expects. No
/// attempt is made to provide a Rust type for COM exceptions.
fn rust_fn_for_com_dispatch_helper(
    context: &'_ mut Context<'_>,
    type_nfo: &ITypeInfo,
    funcdesc: &FUNCDESC,
    fn_name: String,
) -> Result<String, Error> {
    let mut param_types = vec!["&self".to_string()];
    for i in 0..funcdesc.cParams {
        let elemdesc: &mut ELEMDESC =
            unsafe { &mut *funcdesc.lprgelemdescParam.offset(i as isize) };
        param_types.push(format!(
            "param{}: {}",
            i,
            type_bridge::bridge_elem_to_rust_type(context, type_nfo, &elemdesc.tdesc)
        ));
    }

    let return_type = if VARENUM(funcdesc.elemdescFunc.tdesc.vt as i32) == VT_VOID {
        " -> Result<(), HRESULT>".to_string()
    } else {
        format!(
            " -> Result<{}, HRESULT>",
            type_bridge::bridge_elem_to_rust_type(context, type_nfo, &funcdesc.elemdescFunc.tdesc)
        )
    };

    Ok(format!(
        "pub unsafe fn {}({}){}",
        fn_name,
        param_types.join(", "),
        return_type
    ))
}

/// Export a single dispatch property from a type info structure to Rust source
/// code.
pub fn print_type_dispatch_as_rust(
    context: &mut Context<'_>,
    type_nfo: &ITypeInfo,
    fn_index: u32,
) -> Result<String, Error> {
    let mut ret = String::new();

    let funcdesc_raw = unsafe { type_nfo.GetFuncDesc(fn_index)? };
    if funcdesc_raw.is_null() {
        return Err(Error::NoFuncDescForComFn);
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
        FUNC_DISPATCH if funcdesc.invkind != INVOKE_PROPERTYPUTREF => {
            let name = match funcdesc.invkind {
                INVOKE_FUNC => format!("{}", strname),
                INVOKE_PROPERTYGET => format!("{}_get", strname),
                INVOKE_PROPERTYPUT => format!("{}_set", strname),
                _ => unimplemented!(),
            };

            writeln!(
                ret,
                "    {} {{",
                rust_fn_for_com_dispatch_helper(context, type_nfo, funcdesc, name)?
            )?;

            writeln!(ret, "        let mut arg_params = vec![];")?;

            for i in (0..funcdesc.cParams).rev() {
                let elemdesc: &mut ELEMDESC =
                    unsafe { &mut *funcdesc.lprgelemdescParam.offset(i as isize) };

                writeln!(
                    ret,
                    "        {}",
                    dispatch_bridge::generate_param(context, type_nfo, i, &elemdesc.tdesc)?
                )?;
            }

            writeln!(ret, "        let mut disp_params = DISPPARAMS {{")?;
            writeln!(ret, "            rgvarg: arg_params.as_mut_ptr(),")?;
            writeln!(
                ret,
                "            rgdispidNamedArgs: ::std::ptr::null_mut(),"
            )?;
            writeln!(ret, "            cArgs: arg_params.len() as u32,")?;
            writeln!(ret, "            cNamedArgs: 0")?;
            writeln!(ret, "        }};")?;

            if VARENUM(funcdesc.elemdescFunc.tdesc.vt as i32) != VT_VOID {
                writeln!(ret, "        let mut disp_result = VARIANT::default();")?;
                writeln!(
                    ret,
                    "        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
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
                )?;
                writeln!(ret, "        if invoke_result.is_err() {{")?;
                writeln!(ret, "            return Err(invoke_result);")?;
                writeln!(ret, "        }}")?;

                writeln!(
                    ret,
                    "        Ok({})",
                    dispatch_bridge::generate_return(
                        context,
                        type_nfo,
                        &funcdesc.elemdescFunc.tdesc
                    )?
                )?;
            } else {
                writeln!(
                    ret,
                    "        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
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
                )?;
                writeln!(ret, "        if invoke_result.is_err() {{")?;
                writeln!(ret, "            Err(invoke_result)")?;
                writeln!(ret, "        }} else {{")?;
                writeln!(ret, "            Ok(())")?;
                writeln!(ret, "        }}")?;
            }

            writeln!(ret, "    }}")?;
        }
        FUNC_DISPATCH => {
            write!(
                ret,
                "        //TODO: IDispatch helper for {} (invkind {:?})",
                strname, funcdesc.invkind
            )?;
        }
        _ => {}
    }

    unsafe { type_nfo.ReleaseFuncDesc(funcdesc) };

    Ok(ret)
}
