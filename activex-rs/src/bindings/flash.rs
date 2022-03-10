//! Exported from COM metadata to Rust via comexport-rs
//!
//! GUID: D27CDB6B-AE6D-11CF-96B8-444553540000
//! Version: 1.0
//! Locale ID: 0
//! SYSKIND: 3
//! Flags: 8

#![allow(clippy::too_many_arguments)]
#![allow(clippy::upper_case_acronyms)]
#![allow(clippy::missing_safety_doc)]
#![allow(clippy::vec_init_then_push)]
#![allow(non_snake_case)]
#![allow(non_camel_case_types)]

use crate::IDispatch;
use com::interfaces::IUnknown;
use com::sys::GUID;
use com::Interface;
use std::ffi::c_void;
use std::mem::ManuallyDrop;
use windows::core::HRESULT;
use windows::Win32::Foundation::BOOL;
use windows::Win32::System::Com::{
    DISPPARAMS, EXCEPINFO, SAFEARRAY, VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0,
};
use windows::Win32::System::Ole::{DISPATCH_METHOD, VARENUM};

type BSTR = *const u16;
type CY = i64;
type OLE_HANDLE = u32;

/// Shockwave Flash
///
/// GUID: D27CDB6E-AE6D-11CF-96B8-444553540000
/// Interfaces: IShockwaveFlash, _IShockwaveFlashEvents
pub const SHOCKWAVE_FLASH_CLSID: GUID = GUID {
    data1: 0xD27CDB6E,
    data2: 0xAE6D,
    data3: 0x11CF,
    data4: [0x96, 0xB8, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0],
};

/// IFlashObjectInterface Interface
///
/// GUID: D27CDB71-AE6D-11CF-96B8-444553540000
/// Interfaces: IFlashObjectInterface
pub const FLASH_OBJECT_INTERFACE_CLSID: GUID = GUID {
    data1: 0xD27CDB71,
    data2: 0xAE6D,
    data3: 0x11CF,
    data4: [0x96, 0xB8, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0],
};

/// FlashObject Class
///
/// GUID: E0920E11-6B65-4D5D-9C58-B1FC5C07DC43
/// Interfaces: IFlashObject
pub const FLASH_OBJECT_CLSID: GUID = GUID {
    data1: 0xE0920E11,
    data2: 0x6B65,
    data3: 0x4D5D,
    data4: [0x9C, 0x58, 0xB1, 0xFC, 0x5C, 0x7, 0xDC, 0x43],
};

com::interfaces! {
    /// Shockwave Flash
    #[uuid("D27CDB6C-AE6D-11CF-96B8-444553540000")]
    pub unsafe interface IShockwaveFlash: IDispatch {
    }

    /// Event interface for Shockwave Flash
    #[uuid("D27CDB6D-AE6D-11CF-96B8-444553540000")]
    pub unsafe interface _IShockwaveFlashEvents: IDispatch {
    }

    /// IFlashFactory Interface
    #[uuid("D27CDB70-AE6D-11CF-96B8-444553540000")]
    pub unsafe interface IFlashFactory: IUnknown {
    }

    /// IFlashObjectInterface Interface
    #[uuid("D27CDB72-AE6D-11CF-96B8-444553540000")]
    pub unsafe interface IFlashObjectInterface: IDispatchEx {
    }

    #[uuid("A6EF9860-C720-11D0-9337-00A0C90DCAA9")]
    pub unsafe interface IDispatchEx: IDispatch {
        pub unsafe fn GetDispID(&self, param0: BSTR, param1: u32, param2: *mut i32) -> HRESULT;
        pub unsafe fn RemoteInvokeEx(&self, param0: i32, param1: u32, param2: u32, param3: *mut DISPPARAMS, param4: *mut VARIANT, param5: *mut EXCEPINFO, param6: Option<IServiceProvider>, param7: u32, param8: *mut u32, param9: *mut VARIANT) -> HRESULT;
        pub unsafe fn DeleteMemberByName(&self, param0: BSTR, param1: u32) -> HRESULT;
        pub unsafe fn DeleteMemberByDispID(&self, param0: i32) -> HRESULT;
        pub unsafe fn GetMemberProperties(&self, param0: i32, param1: u32, param2: *mut u32) -> HRESULT;
        pub unsafe fn GetMemberName(&self, param0: i32, param1: *mut BSTR) -> HRESULT;
        pub unsafe fn GetNextDispID(&self, param0: u32, param1: i32, param2: *mut i32) -> HRESULT;
        pub unsafe fn GetNameSpaceParent(&self, param0: Option<IUnknown>) -> HRESULT;
    }

    #[uuid("6D5140C1-7436-11CE-8034-00AA006009FA")]
    pub unsafe interface IServiceProvider: IUnknown {
        pub unsafe fn RemoteQueryService(&self, param0: *mut GUID, param1: *mut GUID, param2: Option<IUnknown>) -> HRESULT;
    }

    /// IFlashObject Interface
    #[uuid("86230738-D762-4C50-A2DE-A753E5B1686F")]
    pub unsafe interface IFlashObject: IDispatch {
    }

}

impl IShockwaveFlash {
    pub unsafe fn QueryInterface(
        &self,
        param0: *mut GUID,
        param1: *mut *mut c_void,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_VOID.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param1 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param0 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60000000,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn AddRef(&self) -> Result<u32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60000001,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_UI4
            {
                disp_result.Anonymous.Anonymous.Anonymous.ulVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_UI4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn Release(&self) -> Result<u32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60000002,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_UI4
            {
                disp_result.Anonymous.Anonymous.Anonymous.ulVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_UI4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn GetTypeInfoCount(&self, param0: *mut u32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UINT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { puintVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60010000,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn GetTypeInfo(
        &self,
        param0: u32,
        param1: u32,
        param2: *mut *mut c_void,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_VOID.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param2 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { ulVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UINT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { uintVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60010001,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn GetIDsOfNames(
        &self,
        param0: *mut GUID,
        param1: *mut *mut i8,
        param2: u32,
        param3: u32,
        param4: *mut i32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { plVal: param4 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { ulVal: param3 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UINT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { uintVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I1.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param1 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param0 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60010002,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Invoke(
        &self,
        param0: i32,
        param1: *mut GUID,
        param2: u32,
        param3: u16,
        param4: *mut DISPPARAMS,
        param5: *mut VARIANT,
        param6: *mut EXCEPINFO,
        param7: *mut u32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UINT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { puintVal: param7 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param6 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_VARIANT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { pvarVal: param5 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param4 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI2.0 as u16,
                    Anonymous: VARIANT_0_0_0 { uiVal: param3 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { ulVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param1 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60010003,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn ReadyState_get(&self) -> Result<i32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xFFFFFDF3,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_I4
            {
                disp_result.Anonymous.Anonymous.Anonymous.lVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_I4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn TotalFrames_get(&self) -> Result<i32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x7C,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_I4
            {
                disp_result.Anonymous.Anonymous.Anonymous.lVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_I4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn Playing_get(&self) -> Result<BOOL, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x7D,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BOOL
            {
                BOOL(disp_result.Anonymous.Anonymous.Anonymous.boolVal as i32)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BOOL,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn Playing_set(&self, param0: BOOL) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        boolVal: param0.0 as i16,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x7D,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Quality_get(&self) -> Result<i32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x69,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_INT
            {
                disp_result.Anonymous.Anonymous.Anonymous.intVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_INT,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn Quality_set(&self, param0: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_INT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { intVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x69,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn ScaleMode_get(&self) -> Result<i32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x78,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_INT
            {
                disp_result.Anonymous.Anonymous.Anonymous.intVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_INT,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn ScaleMode_set(&self, param0: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_INT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { intVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x78,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn AlignMode_get(&self) -> Result<i32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x79,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_INT
            {
                disp_result.Anonymous.Anonymous.Anonymous.intVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_INT,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn AlignMode_set(&self, param0: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_INT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { intVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x79,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn BackgroundColor_get(&self) -> Result<i32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x7B,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_I4
            {
                disp_result.Anonymous.Anonymous.Anonymous.lVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_I4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn BackgroundColor_set(&self, param0: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x7B,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Loop_get(&self) -> Result<BOOL, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x6A,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BOOL
            {
                BOOL(disp_result.Anonymous.Anonymous.Anonymous.boolVal as i32)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BOOL,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn Loop_set(&self, param0: BOOL) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        boolVal: param0.0 as i16,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x6A,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Movie_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x66,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn Movie_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x66,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn FrameNum_get(&self) -> Result<i32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x6B,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_I4
            {
                disp_result.Anonymous.Anonymous.Anonymous.lVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_I4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn FrameNum_set(&self, param0: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x6B,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn SetZoomRect(
        &self,
        param0: i32,
        param1: i32,
        param2: i32,
        param3: i32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param3 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x6D,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Zoom(&self, param0: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_INT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { intVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x76,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Pan(&self, param0: i32, param1: i32, param2: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_INT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { intVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x77,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Play(&self) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x70,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Stop(&self) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x71,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Back(&self) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x72,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Forward(&self) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x73,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Rewind(&self) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x74,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn StopPlay(&self) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x7E,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn GotoFrame(&self, param0: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x7F,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn CurrentFrame(&self) -> Result<i32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x80,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_I4
            {
                disp_result.Anonymous.Anonymous.Anonymous.lVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_I4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn IsPlaying(&self) -> Result<BOOL, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x81,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BOOL
            {
                BOOL(disp_result.Anonymous.Anonymous.Anonymous.boolVal as i32)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BOOL,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn PercentLoaded(&self) -> Result<i32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x82,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_I4
            {
                disp_result.Anonymous.Anonymous.Anonymous.lVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_I4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn FrameLoaded(&self, param0: i32) -> Result<BOOL, HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x83,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BOOL
            {
                BOOL(disp_result.Anonymous.Anonymous.Anonymous.boolVal as i32)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BOOL,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn FlashVersion(&self) -> Result<i32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x84,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_I4
            {
                disp_result.Anonymous.Anonymous.Anonymous.lVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_I4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn WMode_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x85,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn WMode_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x85,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn SAlign_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x86,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn SAlign_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x86,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Menu_get(&self) -> Result<BOOL, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x87,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BOOL
            {
                BOOL(disp_result.Anonymous.Anonymous.Anonymous.boolVal as i32)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BOOL,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn Menu_set(&self, param0: BOOL) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        boolVal: param0.0 as i16,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x87,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Base_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x88,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn Base_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x88,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Scale_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x89,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn Scale_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x89,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn DeviceFont_get(&self) -> Result<BOOL, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x8A,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BOOL
            {
                BOOL(disp_result.Anonymous.Anonymous.Anonymous.boolVal as i32)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BOOL,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn DeviceFont_set(&self, param0: BOOL) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        boolVal: param0.0 as i16,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x8A,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn EmbedMovie_get(&self) -> Result<BOOL, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x8B,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BOOL
            {
                BOOL(disp_result.Anonymous.Anonymous.Anonymous.boolVal as i32)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BOOL,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn EmbedMovie_set(&self, param0: BOOL) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        boolVal: param0.0 as i16,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x8B,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn BGColor_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x8C,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn BGColor_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x8C,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Quality2_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x8D,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn Quality2_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x8D,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn LoadMovie(&self, param0: i32, param1: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param1)),
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_INT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { intVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x8E,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn TGotoFrame(&self, param0: BSTR, param1: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x8F,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn TGotoLabel(&self, param0: BSTR, param1: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param1)),
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x90,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn TCurrentFrame(&self, param0: BSTR) -> Result<i32, HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x91,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_I4
            {
                disp_result.Anonymous.Anonymous.Anonymous.lVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_I4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn TCurrentLabel(&self, param0: BSTR) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x92,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn TPlay(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x93,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn TStopPlay(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x94,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn SetVariable(&self, param0: BSTR, param1: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param1)),
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x97,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn GetVariable(&self, param0: BSTR) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x98,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn TSetProperty(
        &self,
        param0: BSTR,
        param1: i32,
        param2: BSTR,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param2)),
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_INT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { intVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x99,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn TGetProperty(&self, param0: BSTR, param1: i32) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_INT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { intVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x9A,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn TCallFrame(&self, param0: BSTR, param1: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_INT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { intVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x9B,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn TCallLabel(&self, param0: BSTR, param1: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param1)),
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x9C,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn TSetPropertyNum(
        &self,
        param0: BSTR,
        param1: i32,
        param2: f64,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_R8.0 as u16,
                    Anonymous: VARIANT_0_0_0 { dblVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_INT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { intVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x9D,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn TGetPropertyNum(&self, param0: BSTR, param1: i32) -> Result<f64, HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_INT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { intVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x9E,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_R8
            {
                disp_result.Anonymous.Anonymous.Anonymous.dblVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_R8,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn TGetPropertyAsNumber(&self, param0: BSTR, param1: i32) -> Result<f64, HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_INT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { intVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xAC,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_R8
            {
                disp_result.Anonymous.Anonymous.Anonymous.dblVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_R8,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn SWRemote_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x9F,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn SWRemote_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x9F,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn FlashVars_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xAA,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn FlashVars_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xAA,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn AllowScriptAccess_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xAB,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn AllowScriptAccess_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xAB,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn MovieData_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xBE,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn MovieData_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xBE,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn InlineData_get(&self) -> Result<Option<IUnknown>, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xBF,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_UNKNOWN
            {
                let com_ptr: Option<IUnknown> =
                    ::std::mem::transmute_copy(&disp_result.Anonymous.Anonymous.Anonymous.punkVal);

                if let Some(com_ptr) = com_ptr.as_ref() {
                    com_ptr.AddRef();
                }

                com_ptr
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_UNKNOWN,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn InlineData_set(&self, param0: Option<IUnknown>) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UNKNOWN.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        punkVal: ::std::mem::transmute(param0.map(|m| m.as_raw())),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xBF,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn SeamlessTabbing_get(&self) -> Result<BOOL, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC0,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BOOL
            {
                BOOL(disp_result.Anonymous.Anonymous.Anonymous.boolVal as i32)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BOOL,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn SeamlessTabbing_set(&self, param0: BOOL) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        boolVal: param0.0 as i16,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC0,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn EnforceLocalSecurity(&self) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC1,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Profile_get(&self) -> Result<BOOL, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC2,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BOOL
            {
                BOOL(disp_result.Anonymous.Anonymous.Anonymous.boolVal as i32)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BOOL,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn Profile_set(&self, param0: BOOL) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        boolVal: param0.0 as i16,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC2,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn ProfileAddress_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC3,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn ProfileAddress_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC3,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn ProfilePort_get(&self) -> Result<i32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC4,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_I4
            {
                disp_result.Anonymous.Anonymous.Anonymous.lVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_I4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn ProfilePort_set(&self, param0: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC4,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn CallFunction(&self, param0: BSTR) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC6,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn SetReturnValue(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC7,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn DisableLocalSecurity(&self) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC8,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn AllowNetworking_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC9,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn AllowNetworking_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC9,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn AllowFullScreen_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xCA,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn AllowFullScreen_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xCA,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn AllowFullScreenInteractive_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x1F5,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn AllowFullScreenInteractive_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x1F5,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn IsDependent_get(&self) -> Result<BOOL, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x1F6,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BOOL
            {
                BOOL(disp_result.Anonymous.Anonymous.Anonymous.boolVal as i32)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BOOL,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn IsDependent_set(&self, param0: BOOL) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        boolVal: param0.0 as i16,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x1F6,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn BrowserZoom_get(&self) -> Result<BSTR, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x1F7,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BSTR
            {
                ::std::mem::transmute(&mut (*disp_result.Anonymous.Anonymous).Anonymous.bstrVal)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BSTR,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn BrowserZoom_set(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x1F7,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn IsTainted_get(&self) -> Result<BOOL, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x1F8,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_BOOL
            {
                BOOL(disp_result.Anonymous.Anonymous.Anonymous.boolVal as i32)
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_BOOL,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn IsTainted_set(&self, param0: BOOL) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        boolVal: param0.0 as i16,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x1F8,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }
}

impl _IShockwaveFlashEvents {
    pub unsafe fn OnReadyStateChange(&self, param0: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xFFFFFD9F,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn OnProgress(&self, param0: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x7A6,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn FSCommand(&self, param0: BSTR, param1: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param1)),
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x96,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn FlashCall(&self, param0: BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0xC5,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }
}

impl IFlashFactory {}

impl IFlashObjectInterface {}

impl IDispatchEx {}

impl IServiceProvider {}

impl IFlashObject {
    pub unsafe fn QueryInterface(
        &self,
        param0: *mut GUID,
        param1: *mut *mut c_void,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_VOID.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param1 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param0 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60000000,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn AddRef(&self) -> Result<u32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60000001,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_UI4
            {
                disp_result.Anonymous.Anonymous.Anonymous.ulVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_UI4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn Release(&self) -> Result<u32, HRESULT> {
        let mut arg_params = vec![];
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let mut disp_result = VARIANT::default();
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60000002,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            &mut disp_result,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            return Err(invoke_result);
        }
        Ok(
            if VARENUM(disp_result.Anonymous.Anonymous.vt as i32)
                == ::windows::Win32::System::Ole::VT_UI4
            {
                disp_result.Anonymous.Anonymous.Anonymous.ulVal
            } else {
                panic!(
                    "Expected value of type {:?}, got {}",
                    ::windows::Win32::System::Ole::VT_UI4,
                    disp_result.Anonymous.Anonymous.vt
                );
            },
        )
    }

    pub unsafe fn GetTypeInfoCount(&self, param0: *mut u32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UINT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { puintVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60010000,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn GetTypeInfo(
        &self,
        param0: u32,
        param1: u32,
        param2: *mut *mut c_void,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_VOID.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param2 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { ulVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UINT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { uintVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60010001,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn GetIDsOfNames(
        &self,
        param0: *mut GUID,
        param1: *mut *mut i8,
        param2: u32,
        param3: u32,
        param4: *mut i32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { plVal: param4 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { ulVal: param3 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UINT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { uintVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I1.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param1 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param0 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60010002,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn Invoke(
        &self,
        param0: i32,
        param1: *mut GUID,
        param2: u32,
        param3: u16,
        param4: *mut DISPPARAMS,
        param5: *mut VARIANT,
        param6: *mut EXCEPINFO,
        param7: *mut u32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UINT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { puintVal: param7 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param6 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_VARIANT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { pvarVal: param5 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param4 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI2.0 as u16,
                    Anonymous: VARIANT_0_0_0 { uiVal: param3 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { ulVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param1 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60010003,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn GetDispID(
        &self,
        param0: BSTR,
        param1: u32,
        param2: *mut i32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { plVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { ulVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60020000,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn RemoteInvokeEx(
        &self,
        param0: i32,
        param1: u32,
        param2: u32,
        param3: *mut DISPPARAMS,
        param4: *mut VARIANT,
        param5: *mut EXCEPINFO,
        param6: Option<IServiceProvider>,
        param7: u32,
        param8: *mut u32,
        param9: *mut VARIANT,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_VARIANT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { pvarVal: param9 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UINT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { puintVal: param8 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UINT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { uintVal: param7 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: ::std::mem::transmute(param6.map(|m| m.as_raw())),
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param5 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_VARIANT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { pvarVal: param4 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_USERDEFINED.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        byref: param3 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { ulVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { ulVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60020001,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn DeleteMemberByName(&self, param0: BSTR, param1: u32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { ulVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(::std::mem::transmute(param0)),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60020002,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn DeleteMemberByDispID(&self, param0: i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60020003,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn GetMemberProperties(
        &self,
        param0: i32,
        param1: u32,
        param2: *mut u32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { pulVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { ulVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60020004,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn GetMemberName(&self, param0: i32, param1: *mut BSTR) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        pbstrVal: ::std::mem::transmute(param1),
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60020005,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn GetNextDispID(
        &self,
        param0: u32,
        param1: i32,
        param2: *mut i32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { plVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param1 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UI4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { ulVal: param0 },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60020006,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }

    pub unsafe fn GetNameSpaceParent(&self, param0: Option<IUnknown>) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_UNKNOWN.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        ppunkVal: ::std::mem::transmute(param0.map(|m| m.as_raw())),
                    },
                    ..Default::default()
                }),
            },
        });
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x60020007,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_METHOD as u16,
            &mut disp_params,
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
            ::std::ptr::null_mut(),
        );
        if invoke_result.is_err() {
            Err(invoke_result)
        } else {
            Ok(())
        }
    }
}
