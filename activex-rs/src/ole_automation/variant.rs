//! Dispatch bridging

use crate::IDispatch;
use com::interfaces::IUnknown;
use com::Interface;
use std::mem::{transmute, transmute_copy, ManuallyDrop};
use windows::core::HRESULT;
use windows::Win32::Foundation::{BOOL, BSTR, CHAR, DISP_E_TYPEMISMATCH, PSTR};
use windows::Win32::System::Com::{CY, VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0};
use windows::Win32::System::Ole::{
    VARENUM, VT_BOOL, VT_BSTR, VT_CY, VT_DISPATCH, VT_I1, VT_I2, VT_I4, VT_I8, VT_INT, VT_R4,
    VT_R8, VT_UI1, VT_UI2, VT_UI4, VT_UI8, VT_UINT, VT_UNKNOWN,
};

/// A type that can be expressed in an `IDispatch` call.
pub trait DynamicType: Sized {
    fn variant_type() -> VARENUM;

    fn into_variant(self) -> VARIANT;

    fn try_from_variant(param: &mut VARIANT) -> Result<Self, HRESULT>;
}

macro_rules! dynamic_type_wrap {
    (bare, $param:ident) => {
        $param
    };
    (newtype, $param:ident) => {
        $param.0
    };
    (bstr, $param:ident) => {
        ManuallyDrop::new(transmute($param))
    };
    (comptr, $param:ident) => {
        transmute($param.map(|m| m.as_raw()))
    };
    (bool, $param:ident) => {
        $param.0 as i16
    };
    (bareptr, $param:ident) => {
        $param
    };
    (bstrptr, $param:ident) => {
        transmute($param)
    };
    (boolptr, $param:ident) => {
        transmute($param) //probably unsound
    };
    (comptrptr, $param:ident) => {
        transmute($param.as_ref().map(|m| m.as_raw()))
    };
}

macro_rules! dynamic_type_unwrap {
    (bare, $param:ident, $varfield:ident, $rs_type:ty) => {
        $param.Anonymous.Anonymous.Anonymous.$varfield
    };
    (newtype, $param:ident, $varfield:ident, $rs_type:ty) => {
        $param.Anonymous.Anonymous.Anonymous.$varfield.into()
    };
    (bstr, $param:ident, $varfield:ident, $rs_type:ty) => {
        transmute(&mut (*$param.Anonymous.Anonymous).Anonymous.$varfield)
    };
    (comptr, $param:ident, $varfield:ident, $rs_type:ty) => {
        transmute_copy::<_, $rs_type>(&$param.Anonymous.Anonymous.Anonymous.$varfield).map(|v| {
            v.AddRef();
            v
        })
    };
    (bool, $param:ident, $varfield:ident, $rs_type:ty) => {
        BOOL($param.Anonymous.Anonymous.Anonymous.$varfield as i32)
    };
    (bareptr, $param:ident, $varfield:ident, $rs_type:ty) => {
        &mut *(*$param.Anonymous.Anonymous).Anonymous.$varfield
    };
    (bstrptr, $param:ident, $varfield:ident, $rs_type:ty) => {
        transmute(&mut (*$param.Anonymous.Anonymous).Anonymous.$varfield)
    };
    (boolptr, $param:ident, $varfield:ident, $rs_type:ty) => {
        &mut *($param.Anonymous.Anonymous.Anonymous.$varfield as *mut $rs_type)
    };

    //TODO: I have no idea how to turn a *mut Option<IUnknown> into a &mut Option<com::IUnknown>
    (comptrptr, $param:ident, $varfield:ident, $rs_type:ty) => {
        transmute_copy::<_, $rs_type>(&$param.Anonymous.Anonymous.Anonymous.$varfield).map(|v| {
            v.AddRef();
            v
        })
    };
}

macro_rules! dynamic_type_impl {
    ($rs_type:ty, $vt:expr, $varfield:ident, $wraptype:tt) => {
        impl DynamicType for $rs_type {
            fn variant_type() -> VARENUM {
                $vt
            }

            fn into_variant(self) -> VARIANT {
                #[allow(unused_unsafe)]
                unsafe {
                    VARIANT {
                        Anonymous: VARIANT_0 {
                            Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                                vt: $vt.0 as u16,
                                Anonymous: VARIANT_0_0_0 {
                                    $varfield: dynamic_type_wrap!($wraptype, self),
                                },
                                ..Default::default()
                            }),
                        },
                    }
                }
            }

            #[allow(unused_mut)]
            fn try_from_variant(mut param: &mut VARIANT) -> Result<Self, HRESULT> {
                unsafe {
                    if VARENUM(param.Anonymous.Anonymous.vt as i32) == $vt {
                        Ok(dynamic_type_unwrap!($wraptype, param, $varfield, $rs_type))
                    } else {
                        Err(DISP_E_TYPEMISMATCH)
                    }
                }
            }
        }
    };
}

#[repr(transparent)]
pub struct WinInt(pub i32);

impl From<i32> for WinInt {
    fn from(param: i32) -> Self {
        WinInt(param)
    }
}

#[repr(transparent)]
pub struct WinUInt(pub u32);

impl From<u32> for WinUInt {
    fn from(param: u32) -> Self {
        WinUInt(param)
    }
}

#[repr(transparent)]
pub struct ComBool(pub i16);

impl From<BOOL> for ComBool {
    fn from(param: BOOL) -> Self {
        Self(param.0 as i16)
    }
}

dynamic_type_impl!(i16, VT_I2, iVal, bare);
dynamic_type_impl!(i32, VT_I4, lVal, bare);
dynamic_type_impl!(f32, VT_R4, fltVal, bare);
dynamic_type_impl!(f64, VT_R8, dblVal, bare);
dynamic_type_impl!(CY, VT_CY, cyVal, bare);
dynamic_type_impl!(BSTR, VT_BSTR, bstrVal, bstr);
dynamic_type_impl!(Option<IDispatch>, VT_DISPATCH, pdispVal, comptr);
dynamic_type_impl!(Option<IUnknown>, VT_UNKNOWN, punkVal, comptr);
dynamic_type_impl!(BOOL, VT_BOOL, boolVal, bool);
dynamic_type_impl!(CHAR, VT_I1, cVal, bare);
dynamic_type_impl!(u8, VT_UI1, bVal, bare);
dynamic_type_impl!(u16, VT_UI2, uiVal, bare);
dynamic_type_impl!(u32, VT_UI4, ulVal, bare);
dynamic_type_impl!(i64, VT_I8, llVal, bare);
dynamic_type_impl!(u64, VT_UI8, ullVal, bare);
dynamic_type_impl!(WinInt, VT_INT, intVal, newtype);
dynamic_type_impl!(WinUInt, VT_UINT, uintVal, newtype);
//TODO: () as void?
dynamic_type_impl!(&mut i16, VT_I2, piVal, bareptr);
dynamic_type_impl!(&mut i32, VT_I4, plVal, bareptr);
dynamic_type_impl!(&mut f32, VT_R4, pfltVal, bareptr);
dynamic_type_impl!(&mut f64, VT_R8, pdblVal, bareptr);
dynamic_type_impl!(&mut CY, VT_CY, pcyVal, bareptr);
dynamic_type_impl!(&mut BSTR, VT_BSTR, pbstrVal, bstrptr);

//TODO: Probably unsound?
//dynamic_type_impl!(&mut Option<IDispatch>, VT_DISPATCH, ppdispVal, comptrptr);
//dynamic_type_impl!(&mut Option<IUnknown>, VT_UNKNOWN, ppdispVal, comptrptr);
dynamic_type_impl!(&mut ComBool, VT_BOOL, pboolVal, boolptr);
dynamic_type_impl!(PSTR, VT_I1, pcVal, bare);
dynamic_type_impl!(&mut u8, VT_UI1, pbVal, bareptr);
dynamic_type_impl!(&mut u16, VT_UI2, puiVal, bareptr);
dynamic_type_impl!(&mut u32, VT_UI4, pulVal, bareptr);
dynamic_type_impl!(&mut i64, VT_I8, pllVal, bareptr);
dynamic_type_impl!(&mut u64, VT_UI8, pullVal, bareptr);
dynamic_type_impl!(&mut WinInt, VT_INT, pintVal, bstrptr);
dynamic_type_impl!(&mut WinUInt, VT_UINT, puintVal, bstrptr);
