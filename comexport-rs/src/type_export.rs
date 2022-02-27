//! Exporters for whole COM types

use crate::context::Context;
use crate::error::Error;
use crate::{fn_export, type_bridge};
use convert_case::{Case, Casing};
use std::fmt::Write;
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::{
    ITypeInfo, ITypeLib, TKIND_COCLASS, TKIND_DISPATCH, TKIND_INTERFACE, TKIND_RECORD, TYPEATTR,
};

/// Generate a doccomment for a COM type
///
/// If this function returns a string with any text, it is guaranteed to also
/// contain at least one newline. Callers should print the returned string
/// directly into the stream without adding a newline.
fn com_type_doccomment(
    context: &Context<'_>,
    type_nfo: &ITypeInfo,
    typeattr: &TYPEATTR,
    strdocstring: BSTR,
) -> Result<String, Error> {
    let mut ret = String::new();

    if !strdocstring.is_empty() {
        writeln!(ret, "/// {}", strdocstring)?;
        writeln!(ret, "///")?;
    }

    writeln!(ret, "/// GUID: {:?}", typeattr.guid)?;

    let mut superinterfaces = vec![];
    for i in 0..typeattr.cImplTypes {
        let href = unsafe { type_nfo.GetRefTypeOfImplType(i as u32)? };
        superinterfaces.push(type_bridge::bridge_usertype_to_rust_type(
            context, type_nfo, href,
        )?);
    }

    if !superinterfaces.is_empty() {
        writeln!(ret, "/// Interfaces: {}", superinterfaces.join(", "))?;
    }

    Ok(ret)
}

/// Export a single class from a type library.
///
/// This should be called outside of a `com::interfaces!` block. It will only
/// print things related to external classes, other types are silently skipped.
pub fn print_type_lib_class_as_rust(
    context: &mut Context<'_>,
    lib: &ITypeLib,
    type_index: u32,
) -> Result<(), Error> {
    let type_nfo = unsafe { lib.GetTypeInfo(type_index)? };
    let typeattr_raw = unsafe { type_nfo.GetTypeAttr()? };
    if typeattr_raw.is_null() {
        return Err(Error::NoTypeAttrForComType);
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
            let doccomment = com_type_doccomment(context, &type_nfo, typeattr, strdocstring)?;
            write!(context.structs, "{}", doccomment)?;

            writeln!(
                context.structs,
                "pub const {}_CLSID: GUID = GUID {{",
                strname.to_string().to_case(Case::UpperSnake)
            )?;
            writeln!(context.structs, "    data1: 0x{:X},", typeattr.guid.data1)?;
            writeln!(context.structs, "    data2: 0x{:X},", typeattr.guid.data2)?;
            writeln!(context.structs, "    data3: 0x{:X},", typeattr.guid.data3)?;
            writeln!(
                context.structs,
                "    data4: [0x{:X}, 0x{:X}, 0x{:X}, 0x{:X}, 0x{:X}, 0x{:X}, 0x{:X}, 0x{:X}],",
                typeattr.guid.data4[0],
                typeattr.guid.data4[1],
                typeattr.guid.data4[2],
                typeattr.guid.data4[3],
                typeattr.guid.data4[4],
                typeattr.guid.data4[5],
                typeattr.guid.data4[6],
                typeattr.guid.data4[7]
            )?;
            writeln!(context.structs, "}};")?;
            writeln!(context.structs)?;
        }
        TKIND_RECORD => {
            let rustname = format!("{}", strname);

            if context.types.type_by_name(&rustname).is_none() {
                context
                    .types
                    .define_generated_bridge(typeattr.guid, rustname);

                let doccomment = com_type_doccomment(context, &type_nfo, typeattr, strdocstring)?;
                write!(context.structs, "{}", doccomment)?;
                writeln!(context.structs, "pub struct {} {{", strname)?;
                writeln!(context.structs, "}}")?;
            }
        }
        TKIND_INTERFACE | TKIND_DISPATCH => {}
        k => {
            writeln!(
                context.structs,
                "//WARN: Unknown type {} of kind {:?}",
                strname, k
            )?;
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
///
/// To export those helper functions, see
/// `print_type_lib_interface_impl_as_rust`.
pub fn print_type_lib_interface_as_rust(
    context: &mut Context<'_>,
    lib: &ITypeLib,
    type_index: u32,
) -> Result<(), Error> {
    let type_nfo = unsafe { lib.GetTypeInfo(type_index)? };
    let typeattr_raw = unsafe { type_nfo.GetTypeAttr()? };
    if typeattr_raw.is_null() {
        return Err(Error::NoTypeAttrForComType);
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

    if context.types.type_by_guid(typeattr.guid).is_none() {
        let rustname = format!("{}", strname);
        context
            .types
            .define_generated_bridge(typeattr.guid, rustname);

        match typeattr.typekind {
            //TODO: Other types of COM types
            TKIND_INTERFACE | TKIND_DISPATCH if typeattr.cImplTypes > 0 => {
                if !strdocstring.is_empty() {
                    writeln!(context.interfaces, "    /// {}", strdocstring)?;
                }

                let mut superinterfaces = vec![];
                for i in 0..typeattr.cImplTypes {
                    let href = unsafe { type_nfo.GetRefTypeOfImplType(i as u32)? };
                    superinterfaces.push(type_bridge::bridge_usertype_to_rust_type(
                        context, &type_nfo, href,
                    )?);
                }
                let superinterfaces = superinterfaces.join(", ");

                writeln!(context.interfaces, "    #[uuid(\"{:?}\")]", typeattr.guid)?;
                writeln!(
                    context.interfaces,
                    "    pub unsafe interface {}: {} {{",
                    strname, superinterfaces
                )?;

                for i in 0..typeattr.cFuncs {
                    let func_decl = fn_export::type_function_as_rust(context, &type_nfo, i as u32)?;
                    if !func_decl.is_empty() {
                        writeln!(context.interfaces, "{}", func_decl)?;
                    }
                }

                writeln!(context.interfaces, "    }}")?;
                writeln!(context.interfaces)?;
            }
            TKIND_INTERFACE | TKIND_DISPATCH => {
                writeln!(
                    context.interfaces,
                    "    // TODO: Bare interface type named {}",
                    strname
                )?;
            }
            _ => {}
        }
    }

    unsafe { type_nfo.ReleaseTypeAttr(typeattr_raw) };

    Ok(())
}

/// Export a single interface's Rust-side helpers from a type library.
///
/// This should be called outside of a `com::interfaces!` block. It will only
/// print things related to interfaces that do not fit inside of that block.
pub fn print_type_lib_interface_impl_as_rust(
    context: &mut Context<'_>,
    lib: &ITypeLib,
    type_index: u32,
) -> Result<(), Error> {
    let type_nfo = unsafe { lib.GetTypeInfo(type_index)? };
    let typeattr_raw = unsafe { type_nfo.GetTypeAttr()? };
    if typeattr_raw.is_null() {
        return Err(Error::NoTypeAttrForComType);
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

    if context.types.type_by_guid(typeattr.guid).is_some() {
        match typeattr.typekind {
            TKIND_INTERFACE | TKIND_DISPATCH => {
                writeln!(context.impls, "impl {} {{", strname)?;

                for i in 0..typeattr.cFuncs {
                    let impl_str =
                        fn_export::print_type_dispatch_as_rust(context, &type_nfo, i as u32)?;
                    writeln!(context.impls, "{}", impl_str)?;
                }

                writeln!(context.impls, "}}")?;
                writeln!(context.impls)?;
            }
            _ => {}
        }
    }

    unsafe { type_nfo.ReleaseTypeAttr(typeattr_raw) };

    Ok(())
}
