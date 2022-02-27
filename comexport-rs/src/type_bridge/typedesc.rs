//! TYPEDESC bridging

use crate::context::Context;
use std::borrow::Cow;
use std::ptr::null_mut;
use windows::core::Error as WinError;
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::{ITypeInfo, TYPEATTR, TYPEDESC};
use windows::Win32::System::Ole::{
    VARENUM, VT_BOOL, VT_BSTR, VT_CY, VT_DISPATCH, VT_HRESULT, VT_I1, VT_I2, VT_I4, VT_I8, VT_INT,
    VT_LPSTR, VT_LPWSTR, VT_PTR, VT_R4, VT_R8, VT_SAFEARRAY, VT_UI1, VT_UI2, VT_UI4, VT_UI8,
    VT_UINT, VT_UNKNOWN, VT_USERDEFINED, VT_VARIANT, VT_VOID,
};

/// Given a type and a referred type ID, print the type it would be if defined
/// in Rust.
///
/// This is sometimes referred to as `HREFTYPE` in Microsoft documentation.
///
/// NOTE: User-defined types are used verbatim. There is currently no machinery
/// to look up or add use statements for user-defined types.
pub fn bridge_usertype_to_rust_type<'a>(
    context: &'a Context<'a>,
    typeinfo: &ITypeInfo,
    href: u32,
) -> Result<Cow<'a, str>, WinError> {
    let mut strname = BSTR::new();

    unsafe {
        let target_type = typeinfo.GetRefTypeInfo(href);
        if target_type.is_err() {
            return Ok(format!(
                "/* unknown user-defined type 0x{:X} (error on reftypeinfo) */",
                href
            )
            .into());
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
            )
            .into());
        }

        let typeattr_raw = target_type.GetTypeAttr()?;
        let typeattr: &mut TYPEATTR = &mut *typeattr_raw;
        if let Some(bridgetype) = context.types.type_by_guid(typeattr.guid) {
            return Ok(bridgetype.rust_name().into());
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
            )
            .into());
        }
    }

    Ok(format!("{}", strname).into())
}

/// Given a type description and the type it came from, print the type it would
/// be if defined in Rust.
///
/// NOTE: User-defined types are used verbatim. There is currently no machinery
/// to look up or add use statements for user-defined types.
pub fn bridge_elem_to_rust_type<'a>(
    context: &'a Context<'a>,
    typeinfo: &ITypeInfo,
    tdesc: &TYPEDESC,
) -> Result<Cow<'a, str>, WinError> {
    Ok(match VARENUM(tdesc.vt as i32) {
        VT_I2 => "i16".into(),
        VT_I4 => "i32".into(),
        VT_R4 => "f32".into(),
        VT_R8 => "f64".into(),
        VT_CY => "CY".into(),
        VT_BSTR => "BSTR".into(),
        VT_DISPATCH => "IDispatch".into(),
        VT_BOOL => "BOOL".into(),
        VT_I1 => "i8".into(),
        VT_UI1 => "u8".into(),
        VT_UI2 => "u16".into(),
        VT_UI4 => "u32".into(),
        VT_I8 => "i64".into(),
        VT_UI8 => "u64".into(),
        VT_INT => "i32".into(),
        VT_UINT => "u32".into(),
        VT_VOID => "c_void".into(),
        VT_HRESULT => "HRESULT".into(),
        VT_UNKNOWN => "IUnknown".into(),
        VT_VARIANT => "VARIANT".into(),
        VT_PTR => {
            let target_type: &TYPEDESC = unsafe { &*tdesc.Anonymous.lptdesc };
            format!(
                "*mut {}",
                bridge_elem_to_rust_type(context, typeinfo, target_type)?
            )
            .into()
        }
        VT_SAFEARRAY => "*mut SAFEARRAY".into(),
        VT_USERDEFINED => {
            let href_type = unsafe { tdesc.Anonymous.hreftype };
            bridge_usertype_to_rust_type(context, typeinfo, href_type)?
        }
        VT_LPSTR => "*mut u8".into(), //NOTE: not an str, Windows only does WTF-8
        VT_LPWSTR => "*mut u16".into(), //TODO: OsStr bridge
        _ => format!("/* unknown type 0x{:X} */", tdesc.vt).into(),
    })
}
