//! *REALLY* Registration-Free COM
//!
//! A series of utility functions intended to make it easier to use COM with
//! executables and libraries directly, bypassing the registry COM mechanism.

use com::sys::GUID;
use com::Interface;
use std::ffi::{c_void, OsStr};
use std::mem::transmute;
use windows::core::{Error as WinError, HRESULT, HSTRING};
use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};

/// Retrieve a COM class object from a DLL, without using registration.
///
/// Useful for weirdly stubborn or deliberately mothballed OLE/ActiveX
/// components, such as Flash Player, which exist in WinSxS but do not register
/// COM classes.
///
/// Typically, you will want your class to be returned with the `IClassFactory`
/// interface. You can then instantiate it with the `create_instance` helper
/// function that the COM create provides.
pub fn get_class_object_by_dll<T: Interface>(path: &OsStr, clsid: &GUID) -> Result<T, WinError> {
    let mut ret = None;
    let dll = unsafe { LoadLibraryW(path).ok()? };
    let get_class_object_proc = unsafe {
        GetProcAddress(dll, "DllGetClassObject")
            .ok_or_else(|| WinError::new(HRESULT(-1), HSTRING::new()))?
    };

    let get_class_object: unsafe extern "system" fn(
        *const GUID,
        *const GUID,
        *mut *mut c_void,
    ) -> HRESULT = unsafe { transmute(get_class_object_proc) };

    unsafe { get_class_object(clsid, &T::IID, &mut ret as *mut _ as _) }.ok()?;

    Ok(ret.unwrap())
}
