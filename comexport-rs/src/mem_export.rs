//! COM struct member type export

use crate::context::Context;
use crate::error::Error;
use crate::type_bridge::bridge_elem_to_rust_type;
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::{ITypeInfo, VARDESC, VAR_PERINSTANCE};

pub fn type_member_as_rust(
    context: &Context<'_>,
    typeinfo: &ITypeInfo,
    var_index: u32,
) -> Result<String, Error> {
    let vardesc_raw = unsafe { typeinfo.GetVarDesc(var_index)? };
    if vardesc_raw.is_null() {
        return Err(Error::NoVarDescForComVar);
    }

    let vardesc: &mut VARDESC = unsafe { &mut *vardesc_raw };
    let typestring = bridge_elem_to_rust_type(context, typeinfo, &vardesc.elemdescVar.tdesc)?;

    let mut strname = BSTR::new();
    let mut strdocstring = BSTR::new();
    let mut whelpcontext = 0;
    let mut strhelpfile = BSTR::new();

    //TODO: This leaks.
    unsafe {
        typeinfo.GetDocumentation(
            vardesc.memid as i32,
            &mut strname,
            &mut strdocstring,
            &mut whelpcontext,
            &mut strhelpfile,
        )?
    };

    let param_name = if strname.is_empty() {
        format!("mem{}", var_index)
    } else {
        format!("{}", strname)
    };

    match vardesc.varkind {
        VAR_PERINSTANCE => Ok(format!("    pub {}: {}", param_name, typestring)),
        k => Ok(format!(
            "    /* unknown varkind {:?} for mem {} */",
            k, var_index
        )),
    }
}
