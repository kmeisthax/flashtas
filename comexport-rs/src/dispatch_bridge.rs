//! Code generators for `IDispatch` helper code generation

use crate::context::Context;
use crate::type_bridge::bridged_hreftype;
use std::borrow::Cow;
use windows::core::Error as WinError;
use windows::Win32::System::Com::{ITypeInfo, TKIND_INTERFACE, TYPEDESC};
use windows::Win32::System::Ole::{
    VARENUM, VT_BOOL, VT_BSTR, VT_CARRAY, VT_CY, VT_DATE, VT_DECIMAL, VT_DISPATCH, VT_EMPTY,
    VT_ERROR, VT_HRESULT, VT_I1, VT_I2, VT_I4, VT_I8, VT_INT, VT_LPSTR, VT_LPWSTR, VT_NULL, VT_PTR,
    VT_R4, VT_R8, VT_SAFEARRAY, VT_UI1, VT_UI2, VT_UI4, VT_UI8, VT_UINT, VT_UNKNOWN,
    VT_USERDEFINED, VT_VARIANT, VT_VOID,
};

/// Retrieve the fully-qualified Rust name of a given VT value.
fn fully_qualified_name_for_vt_value(vt_val: VARENUM) -> &'static str {
    match vt_val {
        VT_EMPTY => "::windows::Win32::System::Ole::VT_EMPTY",
        VT_NULL => "::windows::Win32::System::Ole::VT_NULL",
        VT_I2 => "::windows::Win32::System::Ole::VT_I2",
        VT_I4 => "::windows::Win32::System::Ole::VT_I4",
        VT_R4 => "::windows::Win32::System::Ole::VT_R4",
        VT_R8 => "::windows::Win32::System::Ole::VT_R8",
        VT_CY => "::windows::Win32::System::Ole::VT_CY",
        VT_DATE => "::windows::Win32::System::Ole::VT_DATE",
        VT_BSTR => "::windows::Win32::System::Ole::VT_BSTR",
        VT_DISPATCH => "::windows::Win32::System::Ole::VT_DISPATCH",
        VT_ERROR => "::windows::Win32::System::Ole::VT_ERROR",
        VT_BOOL => "::windows::Win32::System::Ole::VT_BOOL",
        VT_VARIANT => "::windows::Win32::System::Ole::VT_VARIANT",
        VT_UNKNOWN => "::windows::Win32::System::Ole::VT_UNKNOWN",
        VT_DECIMAL => "::windows::Win32::System::Ole::VT_DECIMAL",
        VT_I1 => "::windows::Win32::System::Ole::VT_I1",
        VT_UI1 => "::windows::Win32::System::Ole::VT_UI1",
        VT_UI2 => "::windows::Win32::System::Ole::VT_UI2",
        VT_UI4 => "::windows::Win32::System::Ole::VT_UI4",
        VT_I8 => "::windows::Win32::System::Ole::VT_I8",
        VT_UI8 => "::windows::Win32::System::Ole::VT_UI8",
        VT_INT => "::windows::Win32::System::Ole::VT_INT",
        VT_UINT => "::windows::Win32::System::Ole::VT_UINT",
        VT_VOID => "::windows::Win32::System::Ole::VT_VOID",
        VT_PTR => "::windows::Win32::System::Ole::VT_PTR",
        VT_SAFEARRAY => "::windows::Win32::System::Ole::VT_SAFEARRAY",
        VT_CARRAY => "::windows::Win32::System::Ole::VT_CARRAY",
        VT_USERDEFINED => "::windows::Win32::System::Ole::VT_USERDEFINED",
        VT_LPSTR => "::windows::Win32::System::Ole::VT_LPSTR",
        VT_LPWSTR => "::windows::Win32::System::Ole::VT_LPWSTR",
        VT_HRESULT => "::windows::Win32::System::Ole::HRESULT",
        vt_really_unk => panic!("Unknown VT variant {:?}", vt_really_unk),
    }
}

/// Determine the VT_* name and union field for a given dispatch type passed
/// by pointer.
///
/// The `ITypeInfo` should correspond to the type we are currently generating
/// an `IDispatch` bridge for. The `TYPEDESC` can refer to any type
/// description.
///
/// If this function returns an error, the error value will be a Rust comment
/// that is expected to be printed inline with the source.
fn vt_and_ufield_for_tdesc_ptr<'a>(
    context: &'a Context<'a>,
    typeinfo: &ITypeInfo,
    tdesc: &TYPEDESC,
) -> Result<(&'static str, &'static str, WrapType<'a>), String> {
    let vt_val = VARENUM(tdesc.vt as i32);
    let fq_name = fully_qualified_name_for_vt_value(vt_val);

    let (ufield, wraptype) = match VARENUM(tdesc.vt as i32) {
        VT_I2 => ("piVal", WrapType::Bare),
        VT_I4 => ("plVal", WrapType::Bare),
        VT_R4 => ("pfltVal", WrapType::Bare),
        VT_R8 => ("pdblVal", WrapType::Bare),
        VT_CY => ("pcyVal", WrapType::Bare),
        VT_BSTR => ("pbstrVal", WrapType::BstrPtr),
        VT_BOOL => ("plVal", WrapType::BoolPtr), //TODO: See comments on `BoolPtr`
        VT_I1 => ("pcVal", WrapType::Bare),
        VT_UI1 => ("pbVal", WrapType::Bare),
        VT_UI2 => ("puiVal", WrapType::Bare),
        VT_UI4 => ("pulVal", WrapType::Bare),
        VT_I8 => ("pllVal", WrapType::Bare),
        VT_UI8 => ("pullVal", WrapType::Bare),
        VT_INT => ("pintVal", WrapType::Bare),
        VT_UINT => ("puintVal", WrapType::Bare),
        VT_VOID => ("byref", WrapType::CVoidPtr),

        // TODO: We assume the other side of the dispatch knows to grab the
        // `*mut c_void` and cast it back to the type it was expecting.
        //
        // This is probably not correct; the `VARIANT` structure doesn't seem
        // like it should be able to do `*mut *mut` to things, and I don't
        // expect COM/OLE Automation to know that it needs to transmute to a
        // doubly-indirect pointer.
        VT_PTR => {
            let (vt, _ufield, wraptype) = vt_and_ufield_for_tdesc_ptr(context, typeinfo, unsafe {
                &*tdesc.Anonymous.lptdesc
            })?;

            if matches!(wraptype, WrapType::ComPtr(_)) {
                return Ok((vt, "byref", wraptype));
            } else {
                return Ok((vt, "byref", WrapType::CVoidPtr));
            }
        }
        VT_USERDEFINED => {
            let href = unsafe { tdesc.Anonymous.hreftype };
            let (comtype, name) = bridged_hreftype(context, typeinfo, href)
                .map_err(|e| format!("/* unbridgeable type {} due to {} */", href, e))?;

            if matches!(comtype, TKIND_INTERFACE) {
                ("byref", WrapType::ComPtr(name))
            } else {
                ("byref", WrapType::CVoidPtr)
            }
        }
        VT_DISPATCH => ("byref", WrapType::ComPtr("IDispatch".into())),
        VT_UNKNOWN => ("byref", WrapType::ComPtr("IUnknown".into())),
        VT_LPWSTR | VT_LPSTR => ("byref", WrapType::CVoidPtr),
        VT_VARIANT => ("pvarVal", WrapType::Bare),
        VT_HRESULT => return Err("/* invalid: cannot use VT_HRESULT by ptr */".to_string()),
        _ => return Err(format!("/* pointer to unknown type 0x{:X} */", tdesc.vt)),
    };

    Ok((fq_name, ufield, wraptype))
}

/// Determine the VT_* name and union field for a given dispatch type passed
/// by value.
///
/// The `ITypeInfo` should correspond to the type we are currently generating
/// an `IDispatch` bridge for. The `TYPEDESC` can refer to any type
/// description.
///
/// If this function returns an error, the error value will be a Rust comment
/// that is expected to be printed inline with the source.
fn vt_and_ufield_for_tdesc_value<'a>(
    context: &'a Context<'a>,
    typeinfo: &ITypeInfo,
    tdesc: &TYPEDESC,
) -> Result<(&'static str, &'static str, WrapType<'a>), String> {
    let vt_val = VARENUM(tdesc.vt as i32);
    let fq_name = fully_qualified_name_for_vt_value(vt_val);

    let (ufield, wraptype) = match VARENUM(tdesc.vt as i32) {
        VT_I2 => ("iVal", WrapType::Bare),
        VT_I4 => ("lVal", WrapType::Bare),
        VT_R4 => ("fltVal", WrapType::Bare),
        VT_R8 => ("dblVal", WrapType::Bare),
        VT_CY => ("cyVal", WrapType::Bare),
        VT_BSTR => ("bstrVal", WrapType::Bstr),
        VT_DISPATCH => ("ppdispVal", WrapType::ManuallyDrop),
        VT_BOOL => ("boolVal", WrapType::Bool),
        VT_I1 => ("cVal", WrapType::Bare),
        VT_UI1 => ("bVal", WrapType::Bare),
        VT_UI2 => ("uiVal", WrapType::Bare),
        VT_UI4 => ("ulVal", WrapType::Bare),
        VT_I8 => ("llVal", WrapType::Bare),
        VT_UI8 => ("ullVal", WrapType::Bare),
        VT_INT => ("intVal", WrapType::Bare),
        VT_UINT => ("uintVal", WrapType::Bare),
        VT_VOID => return Err("/* invalid: cannot use VT_VOID by value */".to_string()),
        VT_UNKNOWN => ("ppunkVal", WrapType::ManuallyDrop),
        VT_PTR => {
            return vt_and_ufield_for_tdesc_ptr(context, typeinfo, unsafe {
                &*tdesc.Anonymous.lptdesc
            })
        }
        VT_USERDEFINED => {
            return Err("/* invalid: cannot use VT_USERDEFINED by value */".to_string())
        }
        VT_LPSTR => ("pcVal", WrapType::Bare), //TODO: This is wrong.
        VT_LPWSTR => ("puival", WrapType::Bare), //TODO: This is wrong.
        VT_VARIANT => return Err("/* invalid: cannot use VT_VARIANT by value */".to_string()),
        VT_HRESULT => return Err("/* invalid: cannot use VT_HRESULT by value */".to_string()),
        _ => return Err(format!("/* unknown type 0x{:X} */", tdesc.vt)),
    };

    Ok((fq_name, ufield, wraptype))
}

/// Flag what kind of wrapping the parameter is expected to have.
///
/// `'a` is the lifetime of the current codegen context.
enum WrapType<'a> {
    /// The Rust-side type requires no wrapping.
    Bare,

    /// The Rust-side type must be wrapped in ManuallyDrop.
    ManuallyDrop,

    /// The Rust-side type must be unwrapped from a `windows` `BOOL`.
    Bool,

    /// The Rust-side type must be cast to `i32`.
    ///
    /// TODO: I'm not actually sure if `IDispatch` wants `BOOL` as a `*mut i32`,
    /// or if it wants us to cast to an i16 var (the type of `pBool`) and then
    /// copy that back into the `BOOL` ptr.
    ///
    /// Or it could actually just be stuffing `*mut i32` straight into a
    /// `*mut i16` and just dealing with it 🤡🤡🤡
    BoolPtr,

    /// The Rust-side type is a `*const u16` that must be wrapped into a
    /// `windows` `BSTR` AND a `ManuallyDrop`.
    Bstr,

    /// The Rust-side type is a `*mut *const u16` that must be wrapped into a
    /// `windows` `*mut BSTR`.
    BstrPtr,

    /// The Rust-side type must be cast to `c_void`.
    CVoidPtr,

    /// The Rust-side type implements `com::Interface` and must be unwrapped
    /// with `as_raw` and then cast to `c_void`.
    ///
    /// Unwrapping a `ComPtr` requires wrapping it with `transmute_copy` and
    /// then `AddRef`ing it before returning.
    ComPtr(Cow<'a, str>),
}

/// Generate code that sets a `VARIANT` for a given param in a Dispatch helper
///
/// The `ITypeInfo` should correspond to the type we are currently generating
/// an `IDispatch` bridge for. The `TYPEDESC` should come from the param type
/// we are bridging, and NOT the type this function was defined in.
pub fn generate_param(
    context: &Context<'_>,
    typeinfo: &ITypeInfo,
    param_i: i16,
    tdesc: &TYPEDESC,
) -> Result<String, WinError> {
    let (vt, ufield, wraptype) = match vt_and_ufield_for_tdesc_value(context, typeinfo, tdesc) {
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
        WrapType::BstrPtr => format!("{}: ::std::mem::transmute(param{})", ufield, param_i),
        WrapType::BoolPtr => format!("{}: param{} as *mut i32", ufield, param_i),
        WrapType::CVoidPtr => format!("{}: param{} as *mut c_void", ufield, param_i),
        WrapType::ComPtr(_) => format!(
            "{}: ::std::mem::transmute(param{}.as_raw())",
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
/// The `ITypeInfo` should correspond to the type we are currently generating
/// an `IDispatch` bridge for. The `TYPEDESC` should come from the return type
/// we are bridging, and NOT the type this function was defined in.
///
/// This function assumes that the generated code has a local variable titled
/// `disp_result`, which the function will attempt to access.
pub fn generate_return(
    context: &Context<'_>,
    typeinfo: &ITypeInfo,
    tdesc: &TYPEDESC,
) -> Result<String, WinError> {
    let (vt, ufield, wraptype) = match vt_and_ufield_for_tdesc_value(context, typeinfo, tdesc) {
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
        WrapType::Bstr | WrapType::BstrPtr | WrapType::BoolPtr => format!(
            "::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.{})",
            ufield
        ),
        WrapType::CVoidPtr => format!("disp_result.Anonymous.Anonymous.Anonymous.{}", ufield),
        WrapType::ComPtr(bridged_type) => format!(
            "let com_ptr: {} = ::std::mem::transmute(disp_result.Anonymous.Anonymous.Anonymous.{});
            com_ptr.AddRef();
            com_ptr",
            bridged_type, ufield
        ),
    };

    Ok(format!("if VARENUM(disp_result.Anonymous.Anonymous.vt as i32) == {0} {{
            {1}
        }} else {{
            panic!(\"Expected value of type {{:?}}, got {{}}\", {0}, disp_result.Anonymous.Anonymous.vt);
        }}", vt, wrapper))
}
