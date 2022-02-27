//! COM bindings generator for Rust
//!
//! Produces a Rust module with the type metadata for a given COM type library.
//!
//! Output relies on the `com` and `activex-rs` crates for various bridging
//! utilities.

use com::runtime::init_runtime;
use std::env::args;
use std::ffi::OsStr;
use windows::core::{Error as WinError, HRESULT, HSTRING};
use windows::Win32::System::Com::{ITypeLib, TLIBATTR};
use windows::Win32::System::Ole::LoadTypeLib;

mod fn_export;
mod type_bridge;
mod type_export;

/// Export a type library's metadata as a module doccomment.
fn print_type_lib_doccomment(lib: &ITypeLib) -> Result<(), WinError> {
    let libattr_raw = unsafe { lib.GetLibAttr()? };
    if libattr_raw.is_null() {
        return Err(WinError::new(HRESULT(-1), HSTRING::new()));
    }

    let libattr: &mut TLIBATTR = unsafe { &mut *libattr_raw };

    println!("//! Exported from COM metadata to Rust via comexport-rs");
    println!("//!");
    println!("//! GUID: {:?}", libattr.guid);
    println!(
        "//! Version: {}.{}",
        libattr.wMajorVerNum, libattr.wMinorVerNum
    );
    println!("//! Locale ID: {}", libattr.lcid);
    println!("//! SYSKIND: {}", libattr.syskind.0);
    println!("//! Flags: {}", libattr.wLibFlags);
    println!();

    //TODO: I think this is unsound
    unsafe { lib.ReleaseTLibAttr(libattr_raw) };

    Ok(())
}

fn main() {
    init_runtime().expect("A working COM runtime");

    let path = args().nth(1).expect("Path to .exe, .dll, .tlb, or .ocx");
    let fp_ocx_name = OsStr::new(&path);
    let fp_lib = unsafe { LoadTypeLib(fp_ocx_name).expect("Loaded type lib") };

    unsafe {
        eprintln!("{} has {} entries", path, fp_lib.GetTypeInfoCount());

        print_type_lib_doccomment(&fp_lib).unwrap();

        println!("#![allow(clippy::too_many_arguments)]");
        println!("#![allow(clippy::upper_case_acronyms)]");
        println!("#![allow(clippy::missing_safety_doc)]");
        println!("#![allow(clippy::vec_init_then_push)]");
        println!("#![allow(non_snake_case)]");
        println!();

        //TODO: automatic bridging from user-defined types to `windows`/`com` types
        println!("use crate::IDispatch;");
        println!("use com::interfaces::IUnknown;");
        println!("use com::sys::GUID;");
        println!("use windows::core::HRESULT;");
        println!(
            "use windows::Win32::System::Com::{{
    DISPPARAMS, EXCEPINFO, VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0, SAFEARRAY
}};"
        );
        println!("use windows::Win32::System::Ole::{{DISPATCH_METHOD, VARENUM}};");
        println!("use windows::Win32::Foundation::BOOL;");
        println!("use std::mem::ManuallyDrop;");
        println!("use std::ffi::c_void;");
        println!();
        println!("type BSTR = *const u16;");
        println!("type CY = i64;");
        println!();

        for i in 0..fp_lib.GetTypeInfoCount() {
            type_export::print_type_lib_class_as_rust(&fp_lib, i).unwrap();
        }

        println!("com::interfaces! {{");

        for i in 0..fp_lib.GetTypeInfoCount() {
            type_export::print_type_lib_interface_as_rust(&fp_lib, i).unwrap();
        }

        println!("}}");
        println!();

        for i in 0..fp_lib.GetTypeInfoCount() {
            type_export::print_type_lib_interface_impl_as_rust(&fp_lib, i).unwrap();
        }

        eprintln!("Type definition export complete!");
    }
}
