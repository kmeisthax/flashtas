//! TYPEDESC bridging

use crate::context::Context;
use crate::error::Error;
use std::borrow::Cow;
use std::ptr::null_mut;
use windows::core::Error as WinError;
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::{
    ITypeInfo, SAFEARRAYBOUND, TKIND_INTERFACE, TKIND_RECORD, TYPEATTR, TYPEDESC, TYPEKIND,
};
use windows::Win32::System::Ole::{
    ARRAYDESC, VARENUM, VT_BOOL, VT_BSTR, VT_CARRAY, VT_CY, VT_DISPATCH, VT_HRESULT, VT_I1, VT_I2,
    VT_I4, VT_I8, VT_INT, VT_LPSTR, VT_LPWSTR, VT_PTR, VT_R4, VT_R8, VT_SAFEARRAY, VT_UI1, VT_UI2,
    VT_UI4, VT_UI8, VT_UINT, VT_UNKNOWN, VT_USERDEFINED, VT_VARIANT, VT_VOID,
};

/// Given a type and a referred type ID, yield it's `TYPEKIND` and bridged type
/// name.
pub fn bridged_hreftype<'a>(
    context: &'a Context<'a>,
    typeinfo: &ITypeInfo,
    href: u32,
) -> Result<(TYPEKIND, Cow<'a, str>), WinError> {
    let mut strname = BSTR::new();

    unsafe {
        let target_type = typeinfo.GetRefTypeInfo(href)?;
        let mut target_lib = None;
        let mut target_index = 0;

        target_type.GetContainingTypeLib(&mut target_lib, &mut target_index)?;

        let typeattr_raw = target_type.GetTypeAttr()?;
        let typeattr: &mut TYPEATTR = &mut *typeattr_raw;
        if let Some(bridgetype) = context.types.type_by_guid(typeattr.guid) {
            return Ok((bridgetype.com_type(), bridgetype.rust_name().into()));
        }

        let target_lib = target_lib.unwrap();
        target_lib.GetDocumentation(
            target_index as i32,
            &mut strname,
            null_mut(),
            null_mut(),
            null_mut(),
        )?;

        Ok((typeattr.typekind, format!("{}", strname).into()))
    }
}

/// Given a type and a referred type ID, print the type it would be if defined
/// in Rust.
///
/// This is sometimes referred to as `HREFTYPE` in Microsoft documentation.
///
/// The `TYPEKIND` returned is the typekind that was yielded
///
/// NOTE: User-defined types are used verbatim. There is currently no machinery
/// to look up or add use statements for user-defined types.
pub fn bridge_usertype_to_rust_type<'a>(
    context: &'a Context<'a>,
    typeinfo: &ITypeInfo,
    href: u32,
) -> Cow<'a, str> {
    match bridged_hreftype(context, typeinfo, href) {
        Ok((_, name)) => name,
        Err(e) => format!("/* unbridgeable hreftype {}, because {:?} */", href, e).into(),
    }
}

/// Inner impl function for `bridge_elem_to_rust_type`.
///
/// Output format is the COM typekind, Rust name for the bridged type, and how
/// many `*mut`s should be present. The outer function will compile this into
/// an actual type declaration.
///
/// NOTE: if the returned `TYPEKIND` is `TKIND_INTERFACE`, then the Rust bridge
/// already has a layer of indirection inside itself and you should ignore one
/// pointer.
fn bridge_elem_to_rust_type_impl<'a>(
    context: &'a Context<'a>,
    typeinfo: &ITypeInfo,
    tdesc: &TYPEDESC,
    num_pointers: usize,
) -> Result<(TYPEKIND, Cow<'a, str>, usize), Error> {
    Ok(match VARENUM(tdesc.vt as i32) {
        VT_I2 => (TKIND_RECORD, "i16".into(), num_pointers),
        VT_I4 => (TKIND_RECORD, "i32".into(), num_pointers),
        VT_R4 => (TKIND_RECORD, "f32".into(), num_pointers),
        VT_R8 => (TKIND_RECORD, "f64".into(), num_pointers),
        VT_CY => (TKIND_RECORD, "CY".into(), num_pointers),
        VT_BSTR => (TKIND_RECORD, "BSTR".into(), num_pointers),
        VT_DISPATCH => (TKIND_INTERFACE, "IDispatch".into(), num_pointers),
        VT_BOOL => (TKIND_RECORD, "BOOL".into(), num_pointers),
        VT_I1 => (TKIND_RECORD, "i8".into(), num_pointers),
        VT_UI1 => (TKIND_RECORD, "u8".into(), num_pointers),
        VT_UI2 => (TKIND_RECORD, "u16".into(), num_pointers),
        VT_UI4 => (TKIND_RECORD, "u32".into(), num_pointers),
        VT_I8 => (TKIND_RECORD, "i64".into(), num_pointers),
        VT_UI8 => (TKIND_RECORD, "u64".into(), num_pointers),
        VT_INT => (TKIND_RECORD, "i32".into(), num_pointers),
        VT_UINT => (TKIND_RECORD, "u32".into(), num_pointers),
        VT_VOID => (TKIND_RECORD, "c_void".into(), num_pointers),
        VT_HRESULT => (TKIND_RECORD, "HRESULT".into(), num_pointers),
        VT_UNKNOWN => (TKIND_INTERFACE, "IUnknown".into(), num_pointers),
        VT_VARIANT => (TKIND_RECORD, "VARIANT".into(), num_pointers),
        VT_PTR => {
            let target_type: &TYPEDESC = unsafe { &*tdesc.Anonymous.lptdesc };
            bridge_elem_to_rust_type_impl(context, typeinfo, target_type, num_pointers + 1)?
        }
        VT_SAFEARRAY => (TKIND_RECORD, "*mut SAFEARRAY".into(), num_pointers),
        VT_CARRAY => {
            let arraydesc: &ARRAYDESC = unsafe { &*tdesc.Anonymous.lpadesc };
            let mut elem = bridge_elem_to_rust_type(context, typeinfo, &arraydesc.tdescElem);

            for i in 0..arraydesc.cDims {
                // TODO: When Rust can actually express variable-length structs
                // remove this code
                let bound_base = &arraydesc.rgbounds[0] as *const SAFEARRAYBOUND;
                let bound = unsafe { &*bound_base.offset(i as isize) };
                elem = format!("[{}; {}]", elem, bound.cElements).into();
            }

            (TKIND_RECORD, elem, num_pointers)
        }
        VT_USERDEFINED => {
            let href_type = unsafe { tdesc.Anonymous.hreftype };
            let (com_type, name) = bridged_hreftype(context, typeinfo, href_type)?;

            (com_type, name, num_pointers)
        }
        VT_LPSTR => (TKIND_RECORD, "*mut u8".into(), num_pointers), //NOTE: not an str, Windows only does WTF-8
        VT_LPWSTR => (TKIND_RECORD, "*mut u16".into(), num_pointers), //TODO: OsStr bridge
        u => return Err(u.into()),
    })
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
) -> Cow<'a, str> {
    match bridge_elem_to_rust_type_impl(context, typeinfo, tdesc, 0) {
        Ok((TKIND_INTERFACE, name, num_pointers)) if num_pointers == 0 => {
            format!("/* cannot pass {} by value */", name).into()
        }
        Ok((TKIND_INTERFACE, name, num_pointers)) => {
            format!("{}{}", "*mut ".repeat(num_pointers - 1), name).into()
        }
        Ok((_, name, num_pointers)) if num_pointers == 0 => name,
        Ok((_, name, num_pointers)) => format!("{}{}", "*mut ".repeat(num_pointers), name).into(),
        Err(e) => format!("/* unbridgeable type {}, because {:?} */", tdesc.vt, e).into(),
    }
}
