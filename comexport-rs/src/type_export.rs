//! Exporters for whole COM types

use crate::context::Context;
use crate::error::Error;
use crate::mem_export::type_member_as_rust;
use crate::{fn_export, type_bridge};
use convert_case::{Case, Casing};
use std::fmt::Write;
use windows::core::GUID;
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

    if typeattr.guid != GUID::zeroed() {
        writeln!(ret, "/// GUID: {:?}", typeattr.guid)?;
    }

    let mut superinterfaces = vec![];
    for i in 0..typeattr.cImplTypes {
        let href = unsafe { type_nfo.GetRefTypeOfImplType(i as u32)? };
        superinterfaces.push(type_bridge::bridge_usertype_to_rust_type(
            context, type_nfo, href,
        ));
    }

    if !superinterfaces.is_empty() {
        writeln!(ret, "/// Interfaces: {}", superinterfaces.join(", "))?;
    }

    Ok(ret)
}

/// Export a COM/OLE Automation type from a typelib to the given codegen
/// context.
pub fn gen_typelib_type(
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

    let rustname = format!("{}", strname);
    let already_defined = context.types.type_by_guid(typeattr.guid).is_some()
        || context.types.type_by_name(&rustname).is_some();

    if !already_defined {
        match typeattr.typekind {
            TKIND_COCLASS => {
                let doccomment =
                    com_type_doccomment(context, &type_nfo, typeattr, strdocstring.clone())?;
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
                context
                    .types
                    .define_generated_bridge(typeattr.guid, typeattr.typekind, rustname);

                let doccomment =
                    com_type_doccomment(context, &type_nfo, typeattr, strdocstring.clone())?;
                write!(context.structs, "{}", doccomment)?;
                writeln!(context.structs, "#[repr(C)]")?;
                writeln!(context.structs, "pub struct {} {{", strname)?;

                for i in 0..typeattr.cVars {
                    writeln!(
                        context.structs,
                        "{}{}",
                        type_member_as_rust(context, &type_nfo, i as u32)?,
                        if i + 1 < typeattr.cVars { "," } else { "" }
                    )?;
                }

                writeln!(context.structs, "}}")?;
                writeln!(context.structs)?;
            }
            TKIND_INTERFACE | TKIND_DISPATCH if typeattr.cImplTypes > 0 => {
                context
                    .types
                    .define_generated_bridge(typeattr.guid, typeattr.typekind, rustname);

                if !strdocstring.is_empty() {
                    writeln!(context.interfaces, "    /// {}", strdocstring)?;
                }

                let mut superinterfaces = vec![];
                for i in 0..typeattr.cImplTypes {
                    let href = unsafe { type_nfo.GetRefTypeOfImplType(i as u32)? };
                    superinterfaces.push(type_bridge::bridge_usertype_to_rust_type(
                        context, &type_nfo, href,
                    ));
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

                writeln!(context.impls, "impl {} {{", strname)?;

                for i in 0..typeattr.cFuncs {
                    let impl_str =
                        fn_export::print_type_dispatch_as_rust(context, &type_nfo, i as u32)?;
                    writeln!(context.impls, "{}", impl_str)?;
                }

                writeln!(context.impls, "}}")?;
                writeln!(context.impls)?;
            }
            TKIND_INTERFACE | TKIND_DISPATCH => {
                context
                    .types
                    .define_generated_bridge(typeattr.guid, typeattr.typekind, rustname);

                writeln!(
                    context.interfaces,
                    "    // TODO: Bare interface type named {}",
                    strname
                )?;
            }
            k => {
                writeln!(
                    context.structs,
                    "//WARN: Unknown type {} of kind {:?}",
                    strname, k
                )?;
            }
        }
    }

    unsafe { type_nfo.ReleaseTypeAttr(typeattr_raw) };

    Ok(())
}
