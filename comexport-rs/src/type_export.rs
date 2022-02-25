//! Exporters for whole COM types

use crate::{fn_export, type_bridge};
use convert_case::{Case, Casing};
use windows::core::{Error as WinError, HRESULT, HSTRING};
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::{
    ITypeLib, TKIND_COCLASS, TKIND_DISPATCH, TKIND_INTERFACE, TYPEATTR,
};

/// Export a single class from a type library.
///
/// This should be called outside of a `com::interfaces!` block. It will only
/// print things related to external classes, other types are silently skipped.
pub fn print_type_lib_class_as_rust(lib: &ITypeLib, type_index: u32) -> Result<(), WinError> {
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
        TKIND_COCLASS => {
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
        TKIND_INTERFACE | TKIND_DISPATCH => {}
        k => {
            println!("//WARN: Unknown type {} of kind {:?}", strname, k);
        }
    }

    unsafe { type_nfo.ReleaseTypeAttr(typeattr_raw) };

    Ok(())
}

/// Export a single interface from a type library.
///
/// This should be called within a `com::interfaces!` block. It will only print
/// things related to interfaces, other types are silently skipped.
///
/// Furthermore, only type members that can be expressed within the `com` crate
/// interface block will be printed. Things such as properties and dispatch
/// will be skipped; those must be printed outside the interfaces block as a
/// separate `impl` helper function.
pub fn print_type_lib_interface_as_rust(lib: &ITypeLib, type_index: u32) -> Result<(), WinError> {
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
        TKIND_INTERFACE | TKIND_DISPATCH => {
            if !strdocstring.is_empty() {
                println!("    /// {}", strdocstring);
            }

            let mut superinterfaces = vec![];
            for i in 0..typeattr.cImplTypes {
                let href = unsafe { type_nfo.GetRefTypeOfImplType(i as u32)? };
                superinterfaces.push(type_bridge::bridge_usertype_to_rust_type(&type_nfo, href)?);
            }

            println!("    #[uuid(\"{:?}\")]", typeattr.guid);
            println!(
                "    pub unsafe interface {}: {} {{",
                strname,
                superinterfaces.join(", ")
            );

            for i in 0..typeattr.cFuncs {
                fn_export::print_type_function_as_rust(&type_nfo, i as u32)?;
            }

            println!("    }}");
            println!();
        }
        _ => {}
    }

    unsafe { type_nfo.ReleaseTypeAttr(typeattr_raw) };

    Ok(())
}
