//! DISPPARAMS abstraction

use crate::ole_automation::variant::DynamicType;
use windows::core::HRESULT;
use windows::Win32::Foundation::DISP_E_PARAMNOTOPTIONAL;
use windows::Win32::System::Com::DISPPARAMS;

pub trait DispatchParams {
    /// Retrieve a parameter from a `DISPPARAMS` list and coerce it to the
    /// desired type.
    ///
    /// # Safety
    ///
    /// `DISPPARAMS` must contain pointers to valid memory or `NULL`. `cArgs`
    /// must not exceed the size of the array that `rgvarg` points to.
    unsafe fn get_param<T>(&self, index: usize) -> Result<T, HRESULT>
    where
        T: DynamicType;
}

impl DispatchParams for DISPPARAMS {
    unsafe fn get_param<T>(&self, index: usize) -> Result<T, HRESULT>
    where
        T: DynamicType,
    {
        if self.rgvarg.is_null() || index >= self.cArgs as usize {
            return Err(DISP_E_PARAMNOTOPTIONAL);
        }

        let param_ptr = &mut *self.rgvarg.add(index);
        T::try_from_variant(param_ptr)
    }
}
