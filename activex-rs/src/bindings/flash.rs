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
    DISPPARAMS, EXCEPINFO, VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0,
};
use windows::Win32::System::Ole::{
    DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, VARENUM,
};

type BSTR = *const u16;

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
    pub const QUERY_INTERFACE: u32 = 0x60000000;

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
            Self::QUERY_INTERFACE as i32,
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

    pub const ADD_REF: u32 = 0x60000001;

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
            Self::ADD_REF as i32,
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

    pub const RELEASE: u32 = 0x60000002;

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
            Self::RELEASE as i32,
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

    pub const GET_TYPE_INFO_COUNT: u32 = 0x60010000;

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
            Self::GET_TYPE_INFO_COUNT as i32,
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

    pub const GET_TYPE_INFO: u32 = 0x60010001;

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
            Self::GET_TYPE_INFO as i32,
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

    pub const GET_I_DS_OF_NAMES: u32 = 0x60010002;

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
            Self::GET_I_DS_OF_NAMES as i32,
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

    pub const INVOKE: u32 = 0x60010003;

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
            Self::INVOKE as i32,
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

    pub const READY_STATE_GET: u32 = 0xFFFFFDF3;

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
            Self::READY_STATE_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const TOTAL_FRAMES_GET: u32 = 0x7C;

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
            Self::TOTAL_FRAMES_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const PLAYING_GET: u32 = 0x7D;

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
            Self::PLAYING_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const PLAYING_SET: u32 = 0x7D;

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
            Self::PLAYING_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const QUALITY_GET: u32 = 0x69;

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
            Self::QUALITY_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const QUALITY_SET: u32 = 0x69;

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
            Self::QUALITY_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const SCALE_MODE_GET: u32 = 0x78;

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
            Self::SCALE_MODE_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const SCALE_MODE_SET: u32 = 0x78;

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
            Self::SCALE_MODE_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const ALIGN_MODE_GET: u32 = 0x79;

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
            Self::ALIGN_MODE_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const ALIGN_MODE_SET: u32 = 0x79;

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
            Self::ALIGN_MODE_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const BACKGROUND_COLOR_GET: u32 = 0x7B;

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
            Self::BACKGROUND_COLOR_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const BACKGROUND_COLOR_SET: u32 = 0x7B;

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
            Self::BACKGROUND_COLOR_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const LOOP_GET: u32 = 0x6A;

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
            Self::LOOP_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const LOOP_SET: u32 = 0x6A;

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
            Self::LOOP_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const MOVIE_GET: u32 = 0x66;

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
            Self::MOVIE_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const MOVIE_SET: u32 = 0x66;

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
            Self::MOVIE_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const FRAME_NUM_GET: u32 = 0x6B;

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
            Self::FRAME_NUM_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const FRAME_NUM_SET: u32 = 0x6B;

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
            Self::FRAME_NUM_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const SET_ZOOM_RECT: u32 = 0x6D;

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
            Self::SET_ZOOM_RECT as i32,
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

    pub const ZOOM: u32 = 0x76;

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
            Self::ZOOM as i32,
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

    pub const PAN: u32 = 0x77;

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
            Self::PAN as i32,
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

    pub const PLAY: u32 = 0x70;

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
            Self::PLAY as i32,
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

    pub const STOP: u32 = 0x71;

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
            Self::STOP as i32,
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

    pub const BACK: u32 = 0x72;

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
            Self::BACK as i32,
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

    pub const FORWARD: u32 = 0x73;

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
            Self::FORWARD as i32,
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

    pub const REWIND: u32 = 0x74;

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
            Self::REWIND as i32,
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

    pub const STOP_PLAY: u32 = 0x7E;

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
            Self::STOP_PLAY as i32,
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

    pub const GOTO_FRAME: u32 = 0x7F;

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
            Self::GOTO_FRAME as i32,
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

    pub const CURRENT_FRAME: u32 = 0x80;

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
            Self::CURRENT_FRAME as i32,
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

    pub const IS_PLAYING: u32 = 0x81;

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
            Self::IS_PLAYING as i32,
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

    pub const PERCENT_LOADED: u32 = 0x82;

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
            Self::PERCENT_LOADED as i32,
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

    pub const FRAME_LOADED: u32 = 0x83;

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
            Self::FRAME_LOADED as i32,
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

    pub const FLASH_VERSION: u32 = 0x84;

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
            Self::FLASH_VERSION as i32,
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

    pub const W_MODE_GET: u32 = 0x85;

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
            Self::W_MODE_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const W_MODE_SET: u32 = 0x85;

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
            Self::W_MODE_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const S_ALIGN_GET: u32 = 0x86;

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
            Self::S_ALIGN_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const S_ALIGN_SET: u32 = 0x86;

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
            Self::S_ALIGN_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const MENU_GET: u32 = 0x87;

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
            Self::MENU_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const MENU_SET: u32 = 0x87;

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
            Self::MENU_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const BASE_GET: u32 = 0x88;

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
            Self::BASE_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const BASE_SET: u32 = 0x88;

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
            Self::BASE_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const SCALE_GET: u32 = 0x89;

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
            Self::SCALE_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const SCALE_SET: u32 = 0x89;

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
            Self::SCALE_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const DEVICE_FONT_GET: u32 = 0x8A;

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
            Self::DEVICE_FONT_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const DEVICE_FONT_SET: u32 = 0x8A;

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
            Self::DEVICE_FONT_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const EMBED_MOVIE_GET: u32 = 0x8B;

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
            Self::EMBED_MOVIE_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const EMBED_MOVIE_SET: u32 = 0x8B;

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
            Self::EMBED_MOVIE_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const BG_COLOR_GET: u32 = 0x8C;

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
            Self::BG_COLOR_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const BG_COLOR_SET: u32 = 0x8C;

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
            Self::BG_COLOR_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const QUALITY_2_GET: u32 = 0x8D;

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
            Self::QUALITY_2_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const QUALITY_2_SET: u32 = 0x8D;

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
            Self::QUALITY_2_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const LOAD_MOVIE: u32 = 0x8E;

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
            Self::LOAD_MOVIE as i32,
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

    pub const T_GOTO_FRAME: u32 = 0x8F;

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
            Self::T_GOTO_FRAME as i32,
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

    pub const T_GOTO_LABEL: u32 = 0x90;

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
            Self::T_GOTO_LABEL as i32,
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

    pub const T_CURRENT_FRAME: u32 = 0x91;

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
            Self::T_CURRENT_FRAME as i32,
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

    pub const T_CURRENT_LABEL: u32 = 0x92;

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
            Self::T_CURRENT_LABEL as i32,
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

    pub const T_PLAY: u32 = 0x93;

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
            Self::T_PLAY as i32,
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

    pub const T_STOP_PLAY: u32 = 0x94;

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
            Self::T_STOP_PLAY as i32,
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

    pub const SET_VARIABLE: u32 = 0x97;

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
            Self::SET_VARIABLE as i32,
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

    pub const GET_VARIABLE: u32 = 0x98;

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
            Self::GET_VARIABLE as i32,
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

    pub const T_SET_PROPERTY: u32 = 0x99;

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
            Self::T_SET_PROPERTY as i32,
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

    pub const T_GET_PROPERTY: u32 = 0x9A;

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
            Self::T_GET_PROPERTY as i32,
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

    pub const T_CALL_FRAME: u32 = 0x9B;

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
            Self::T_CALL_FRAME as i32,
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

    pub const T_CALL_LABEL: u32 = 0x9C;

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
            Self::T_CALL_LABEL as i32,
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

    pub const T_SET_PROPERTY_NUM: u32 = 0x9D;

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
            Self::T_SET_PROPERTY_NUM as i32,
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

    pub const T_GET_PROPERTY_NUM: u32 = 0x9E;

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
            Self::T_GET_PROPERTY_NUM as i32,
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

    pub const T_GET_PROPERTY_AS_NUMBER: u32 = 0xAC;

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
            Self::T_GET_PROPERTY_AS_NUMBER as i32,
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

    pub const SW_REMOTE_GET: u32 = 0x9F;

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
            Self::SW_REMOTE_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const SW_REMOTE_SET: u32 = 0x9F;

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
            Self::SW_REMOTE_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const FLASH_VARS_GET: u32 = 0xAA;

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
            Self::FLASH_VARS_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const FLASH_VARS_SET: u32 = 0xAA;

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
            Self::FLASH_VARS_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const ALLOW_SCRIPT_ACCESS_GET: u32 = 0xAB;

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
            Self::ALLOW_SCRIPT_ACCESS_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const ALLOW_SCRIPT_ACCESS_SET: u32 = 0xAB;

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
            Self::ALLOW_SCRIPT_ACCESS_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const MOVIE_DATA_GET: u32 = 0xBE;

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
            Self::MOVIE_DATA_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const MOVIE_DATA_SET: u32 = 0xBE;

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
            Self::MOVIE_DATA_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const INLINE_DATA_GET: u32 = 0xBF;

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
            Self::INLINE_DATA_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const INLINE_DATA_SET: u32 = 0xBF;

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
            Self::INLINE_DATA_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const SEAMLESS_TABBING_GET: u32 = 0xC0;

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
            Self::SEAMLESS_TABBING_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const SEAMLESS_TABBING_SET: u32 = 0xC0;

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
            Self::SEAMLESS_TABBING_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const ENFORCE_LOCAL_SECURITY: u32 = 0xC1;

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
            Self::ENFORCE_LOCAL_SECURITY as i32,
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

    pub const PROFILE_GET: u32 = 0xC2;

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
            Self::PROFILE_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const PROFILE_SET: u32 = 0xC2;

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
            Self::PROFILE_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const PROFILE_ADDRESS_GET: u32 = 0xC3;

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
            Self::PROFILE_ADDRESS_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const PROFILE_ADDRESS_SET: u32 = 0xC3;

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
            Self::PROFILE_ADDRESS_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const PROFILE_PORT_GET: u32 = 0xC4;

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
            Self::PROFILE_PORT_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const PROFILE_PORT_SET: u32 = 0xC4;

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
            Self::PROFILE_PORT_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const CALL_FUNCTION: u32 = 0xC6;

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
            Self::CALL_FUNCTION as i32,
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

    pub const SET_RETURN_VALUE: u32 = 0xC7;

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
            Self::SET_RETURN_VALUE as i32,
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

    pub const DISABLE_LOCAL_SECURITY: u32 = 0xC8;

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
            Self::DISABLE_LOCAL_SECURITY as i32,
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

    pub const ALLOW_NETWORKING_GET: u32 = 0xC9;

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
            Self::ALLOW_NETWORKING_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const ALLOW_NETWORKING_SET: u32 = 0xC9;

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
            Self::ALLOW_NETWORKING_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const ALLOW_FULL_SCREEN_GET: u32 = 0xCA;

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
            Self::ALLOW_FULL_SCREEN_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const ALLOW_FULL_SCREEN_SET: u32 = 0xCA;

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
            Self::ALLOW_FULL_SCREEN_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const ALLOW_FULL_SCREEN_INTERACTIVE_GET: u32 = 0x1F5;

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
            Self::ALLOW_FULL_SCREEN_INTERACTIVE_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const ALLOW_FULL_SCREEN_INTERACTIVE_SET: u32 = 0x1F5;

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
            Self::ALLOW_FULL_SCREEN_INTERACTIVE_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const IS_DEPENDENT_GET: u32 = 0x1F6;

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
            Self::IS_DEPENDENT_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const IS_DEPENDENT_SET: u32 = 0x1F6;

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
            Self::IS_DEPENDENT_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const BROWSER_ZOOM_GET: u32 = 0x1F7;

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
            Self::BROWSER_ZOOM_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const BROWSER_ZOOM_SET: u32 = 0x1F7;

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
            Self::BROWSER_ZOOM_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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

    pub const IS_TAINTED_GET: u32 = 0x1F8;

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
            Self::IS_TAINTED_GET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYGET as u16,
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

    pub const IS_TAINTED_SET: u32 = 0x1F8;

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
            Self::IS_TAINTED_SET as i32,
            &mut GUID {
                data1: 0,
                data2: 0,
                data3: 0,
                data4: [0; 8],
            },
            0,
            DISPATCH_PROPERTYPUT as u16,
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
    pub const ON_READY_STATE_CHANGE: u32 = 0xFFFFFD9F;

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
            Self::ON_READY_STATE_CHANGE as i32,
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

    pub const ON_PROGRESS: u32 = 0x7A6;

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
            Self::ON_PROGRESS as i32,
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

    pub const FS_COMMAND: u32 = 0x96;

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
            Self::FS_COMMAND as i32,
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

    pub const FLASH_CALL: u32 = 0xC5;

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
            Self::FLASH_CALL as i32,
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
    pub const QUERY_INTERFACE: u32 = 0x60000000;

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
            Self::QUERY_INTERFACE as i32,
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

    pub const ADD_REF: u32 = 0x60000001;

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
            Self::ADD_REF as i32,
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

    pub const RELEASE: u32 = 0x60000002;

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
            Self::RELEASE as i32,
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

    pub const GET_TYPE_INFO_COUNT: u32 = 0x60010000;

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
            Self::GET_TYPE_INFO_COUNT as i32,
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

    pub const GET_TYPE_INFO: u32 = 0x60010001;

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
            Self::GET_TYPE_INFO as i32,
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

    pub const GET_I_DS_OF_NAMES: u32 = 0x60010002;

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
            Self::GET_I_DS_OF_NAMES as i32,
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

    pub const INVOKE: u32 = 0x60010003;

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
            Self::INVOKE as i32,
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

    pub const GET_DISP_ID: u32 = 0x60020000;

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
            Self::GET_DISP_ID as i32,
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

    pub const REMOTE_INVOKE_EX: u32 = 0x60020001;

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
            Self::REMOTE_INVOKE_EX as i32,
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

    pub const DELETE_MEMBER_BY_NAME: u32 = 0x60020002;

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
            Self::DELETE_MEMBER_BY_NAME as i32,
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

    pub const DELETE_MEMBER_BY_DISP_ID: u32 = 0x60020003;

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
            Self::DELETE_MEMBER_BY_DISP_ID as i32,
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

    pub const GET_MEMBER_PROPERTIES: u32 = 0x60020004;

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
            Self::GET_MEMBER_PROPERTIES as i32,
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

    pub const GET_MEMBER_NAME: u32 = 0x60020005;

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
            Self::GET_MEMBER_NAME as i32,
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

    pub const GET_NEXT_DISP_ID: u32 = 0x60020006;

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
            Self::GET_NEXT_DISP_ID as i32,
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

    pub const GET_NAME_SPACE_PARENT: u32 = 0x60020007;

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
            Self::GET_NAME_SPACE_PARENT as i32,
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
