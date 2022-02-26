//! Rust/COM type bridging helpers
//!
//! Functions in this package are responsible for taking a COM type and picking
//! a suitable Rust equivalent that matches.

use std::ptr::null_mut;
use windows::core::Error as WinError;
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::{ITypeInfo, TYPEDESC};
use windows::Win32::System::Ole::{
    VARENUM, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_HRESULT, VT_I1, VT_I2, VT_I4, VT_I8, VT_INT, VT_PTR,
    VT_R4, VT_R8, VT_UI1, VT_UI2, VT_UI4, VT_UI8, VT_UINT, VT_UNKNOWN, VT_USERDEFINED, VT_VARIANT,
    VT_VOID,
};

/// Given a type and a referred type ID, print the type it would be if defined
/// in Rust.
///
/// This is sometimes referred to as `HREFTYPE` in Microsoft documentation.
///
/// NOTE: User-defined types are used verbatim. There is currently no machinery
/// to look up or add use statements for user-defined types.
pub fn bridge_usertype_to_rust_type(typeinfo: &ITypeInfo, href: u32) -> Result<String, WinError> {
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
pub fn bridge_elem_to_rust_type(
    typeinfo: &ITypeInfo,
    tdesc: &TYPEDESC,
) -> Result<String, WinError> {
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
        VT_INT => "i32".to_string(),
        VT_UINT => "u32".to_string(),
        VT_VOID => "c_void".to_string(),
        VT_HRESULT => "HRESULT".to_string(),
        VT_UNKNOWN => "IUnknown".to_string(),
        VT_VARIANT => "VARIANT".to_string(),
        VT_PTR => {
            let target_type: &TYPEDESC = unsafe { &*tdesc.Anonymous.lptdesc };
            format!("*mut {}", bridge_elem_to_rust_type(typeinfo, target_type)?)
        }
        VT_USERDEFINED => {
            let href_type = unsafe { tdesc.Anonymous.hreftype };
            bridge_usertype_to_rust_type(typeinfo, href_type)?
        }
        _ => format!("/* unknown type 0x{:X} */", tdesc.vt),
    })
}

/// Flag what kind of wrapping the parameter is expected to have.
enum WrapType {
    /// The Rust-side type requires no wrapping.
    Bare,

    /// The Rust-side type must be wrapped in ManuallyDrop.
    ManuallyDrop,

    /// The Rust-side type must be unwrapped from a `windows` `BOOL`.
    Bool,

    /// The Rust-side type is a `*const u16` that must be wrapped into a
    /// `windows` `BSTR` AND a `ManuallyDrop`.
    Bstr,
}

/// Determine the VT_* value and union field for a given dispatch type.
///
/// If this function returns an error, the error value will be a Rust comment
/// that is expected to be printed inline with the source.
///
/// The boolean parameter indicates if the field needs to be wrapped in
/// `ManuallyDrop`.
fn vt_and_ufield_for_tdesc(tdesc: &TYPEDESC) -> Result<(&str, &str, WrapType), String> {
    Ok(match VARENUM(tdesc.vt as i32) {
        VT_I2 => (
            "::windows::Win32::System::Ole::VT_I2",
            "iVal",
            WrapType::Bare,
        ),
        VT_I4 => (
            "::windows::Win32::System::Ole::VT_I4",
            "lVal",
            WrapType::Bare,
        ),
        VT_R4 => (
            "::windows::Win32::System::Ole::VT_R4",
            "fltVal",
            WrapType::Bare,
        ),
        VT_R8 => (
            "::windows::Win32::System::Ole::VT_R8",
            "dblVal",
            WrapType::Bare,
        ),
        VT_BSTR => (
            "::windows::Win32::System::Ole::VT_BSTR",
            "bstrVal",
            WrapType::Bstr,
        ),
        VT_DISPATCH => (
            "::windows::Win32::System::Ole::VT_DISPATCH",
            "ppdispVal",
            WrapType::ManuallyDrop,
        ),
        VT_BOOL => (
            "::windows::Win32::System::Ole::VT_BOOL",
            "boolVal",
            WrapType::Bool,
        ),
        VT_I1 => (
            "::windows::Win32::System::Ole::VT_I1",
            "cVal",
            WrapType::Bare,
        ),
        VT_UI1 => (
            "::windows::Win32::System::Ole::VT_UI1",
            "bVal",
            WrapType::Bare,
        ),
        VT_UI2 => (
            "::windows::Win32::System::Ole::VT_UI2",
            "uiVal",
            WrapType::Bare,
        ),
        VT_UI4 => (
            "::windows::Win32::System::Ole::VT_UI4",
            "ulVal",
            WrapType::Bare,
        ),
        VT_I8 => (
            "::windows::Win32::System::Ole::VT_I8",
            "llVal",
            WrapType::Bare,
        ),
        VT_UI8 => (
            "::windows::Win32::System::Ole::VT_UI8",
            "ullVal",
            WrapType::Bare,
        ),
        VT_INT => (
            "::windows::Win32::System::Ole::VT_INT",
            "intVal",
            WrapType::Bare,
        ),
        VT_UINT => (
            "::windows::Win32::System::Ole::VT_UINT",
            "uintVal",
            WrapType::Bare,
        ),
        VT_VOID => (
            "::windows::Win32::System::Ole::VT_VOID",
            "byref",
            WrapType::Bare,
        ),
        VT_HRESULT => (
            "::windows::Win32::System::Ole::VT_HRESULT",
            "lVal",
            WrapType::ManuallyDrop,
        ),
        VT_UNKNOWN => (
            "::windows::Win32::System::Ole::VT_UNKNOWN",
            "ppunkVal",
            WrapType::ManuallyDrop,
        ),
        VT_PTR => return Err("/* invalid: cannot use VT_PTR in IDispatch */".to_string()),
        VT_USERDEFINED | VT_VARIANT => {
            return Err("/* invalid: cannot use VT_VARIANT without byref */".to_string())
        }
        _ => return Err(format!("/* unknown type 0x{:X} */", tdesc.vt)),
    })
}

/// Generate code that sets a `VARIANT` for a given param in a Dispatch helper
pub fn generate_dispatch_param(param_i: i16, tdesc: &TYPEDESC) -> Result<String, WinError> {
    let (vt, ufield, wraptype) = match vt_and_ufield_for_tdesc(tdesc) {
        Ok((vt, ufield, wraptype)) => (vt, ufield, wraptype),
        Err(e) => return Ok(e),
    };

    let wrapper = match wraptype {
        WrapType::Bare => format!("{}: param{}", ufield, param_i),
        WrapType::ManuallyDrop => format!("{}: ManuallyDrop::new(param{})", ufield, param_i),
        WrapType::Bool => format!("{}: param{}.0", ufield, param_i),
        WrapType::Bstr => format!(
            "{}: ManuallyDrop::new(::std::mem::transmute(param{}))",
            ufield, param_i
        ),
    };

    Ok(format!(
        "arg_params.push(VARIANT {{
            Anonymous: VARIANT_0 {{
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {{
                    vt: {0}.0 as u16,
                    Anonymous: VARIANT_0_0_0 {{
                        {1}
                    }},
                    ..Default::default()
                }})
            }}
        }});",
        vt, wrapper
    ))
}

/// Generate code that unwraps a `VARIANT` back into a Rust type.
///
/// This function assumes that the generated code has a local variable titled
/// `disp_result`, which the function will attempt to access.
pub fn generate_dispatch_return(tdesc: &TYPEDESC) -> Result<String, WinError> {
    let (vt, ufield, wraptype) = match vt_and_ufield_for_tdesc(tdesc) {
        Ok((vt, ufield, wraptype)) => (vt, ufield, wraptype),
        Err(e) => return Ok(e),
    };

    let wrapper = match wraptype {
        WrapType::Bare => format!("disp_result.Anonymous.Anonymous.Anonymous.{}", ufield),
        WrapType::ManuallyDrop => format!("disp_result.Anonymous.Anonymous.Anonymous.{}.0", ufield),
        WrapType::Bool => format!(
            "BOOL(disp_result.Anonymous.Anonymous.Anonymous.{} as i32)",
            ufield
        ),
        WrapType::Bstr => format!(
            "::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.{})",
            ufield
        ),
    };

    Ok(format!("if VARENUM(disp_result.Anonymous.Anonymous.vt as i32) == {0} {{
            {1}
        }} else {{
            panic!(\"Expected value of type {{:?}}, got {{}}\", {0}, disp_result.Anonymous.Anonymous.vt);
        }}", vt, wrapper))
}
