//! Exported from COM metadata to Rust via comexport-rs
//!
//! GUID: 00020430-0000-0000-C000-000000000046
//! Version: 2.0
//! Locale ID: 0
//! SYSKIND: 1
//! Flags: 8

#![allow(clippy::too_many_arguments)]
#![allow(clippy::upper_case_acronyms)]

use com::interfaces::IUnknown;
use com::sys::GUID;
use std::ffi::c_void;
use windows::core::HRESULT;
use windows::Win32::System::Com::{DISPPARAMS, EXCEPINFO, VARIANT};

com::interfaces! {
    #[uuid("00020400-0000-0000-C000-000000000046")]
    pub unsafe interface IDispatch: IUnknown {
        pub unsafe fn GetTypeInfoCount(&self, param0: *mut usize) -> HRESULT;
        pub unsafe fn GetTypeInfo(&self, param0: usize, param1: u32, param2: *mut *mut c_void) -> HRESULT;
        pub unsafe fn GetIDsOfNames(&self, param0: *mut GUID, param1: *mut *mut i8, param2: usize, param3: u32, param4: *mut i32) -> HRESULT;
        pub unsafe fn Invoke(&self, param0: i32, param1: *mut GUID, param2: u32, param3: u16, param4: *mut DISPPARAMS, param5: *mut VARIANT, param6: *mut EXCEPINFO, param7: *mut usize) -> HRESULT;
    }
}
