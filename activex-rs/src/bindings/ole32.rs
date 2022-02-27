//! Exported from COM metadata to Rust via comexport-rs
//!
//! GUID: 5A2B9220-BF07-11E6-9598-0800200C9A66
//! Version: 1.0
//! Locale ID: 0
//! SYSKIND: 1
//! Flags: 8

#![allow(clippy::too_many_arguments)]
#![allow(clippy::upper_case_acronyms)]
#![allow(clippy::missing_safety_doc)]
#![allow(clippy::vec_init_then_push)]
#![allow(non_snake_case)]

use crate::IDispatch;
use com::interfaces::IUnknown;
use com::sys::GUID;
use std::ffi::c_void;
use std::mem::ManuallyDrop;
use windows::core::HRESULT;
use windows::Win32::Foundation::BOOL;
use windows::Win32::System::Com::{
    CY, DISPPARAMS, EXCEPINFO, SAFEARRAY, VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0,
};
use windows::Win32::System::Ole::{DISPATCH_METHOD, VARENUM};

type BSTR = *const u16;

//WARN: Unknown type OLERECT of kind TYPEKIND(1)
//WARN: Unknown type OLEPOINT of kind TYPEKIND(1)
//WARN: Unknown type OLEACCELMSG of kind TYPEKIND(1)
//WARN: Unknown type OLESIZE of kind TYPEKIND(1)
//WARN: Unknown type OLEINPLACEFRAMEINFO of kind TYPEKIND(1)
//WARN: Unknown type OLECONTROLINFO of kind TYPEKIND(1)
//WARN: Unknown type OLECLSID of kind TYPEKIND(1)
//WARN: Unknown type OLECAUUID of kind TYPEKIND(1)
//WARN: Unknown type OLECALPOLESTR of kind TYPEKIND(1)
//WARN: Unknown type OLECADWORD of kind TYPEKIND(1)
com::interfaces! {
    // TODO: Bare interface typp named IUnknownUnrestricted
    #[uuid("00000114-0000-0000-C000-000000000046")]
    pub unsafe interface IOleWindow: IUnknown {
        pub unsafe fn GetWindow(&self, param0: *mut i32) -> HRESULT;
        pub unsafe fn ContextSensitiveHelp(&self, param0: i32) -> HRESULT;
    }

    #[uuid("00000118-0000-0000-C000-000000000046")]
    pub unsafe interface IOleClientSite: IUnknown {
    }

    #[uuid("00000112-0000-0000-C000-000000000046")]
    pub unsafe interface IOleObject: IUnknown {
        pub unsafe fn SetClientSite(&self, param0: *mut IOleClientSite) -> HRESULT;
        pub unsafe fn GetClientSite(&self, param0: *mut *mut IOleClientSite) -> HRESULT;
        pub unsafe fn SetHostNames(&self, param0: *mut u16, param1: *mut u16) -> HRESULT;
        pub unsafe fn Close(&self, param0: i32) -> HRESULT;
        pub unsafe fn SetMoniker(&self, param0: i32, param1: i32) -> HRESULT;
        pub unsafe fn GetMoniker(&self, param0: i32, param1: i32, param2: *mut i32) -> HRESULT;
        pub unsafe fn InitFromData(&self, param0: i32, param1: i32, param2: i32) -> HRESULT;
        pub unsafe fn GetClipboardData(&self, param0: i32, param1: *mut i32) -> HRESULT;
        pub unsafe fn DoVerb(&self, param0: i32, param1: i32, param2: *mut IOleClientSite, param3: i32, param4: i32, param5: i32) -> i32;
    }

    #[uuid("00000113-0000-0000-C000-000000000046")]
    pub unsafe interface IOleInPlaceObject: IOleWindow {
        pub unsafe fn InPlaceDeactivate(&self) -> HRESULT;
        pub unsafe fn UIDeactivate(&self) -> HRESULT;
        pub unsafe fn SetObjectRects(&self, param0: i32, param1: i32) -> HRESULT;
        pub unsafe fn ReactivateAndUndo(&self) -> HRESULT;
    }

    #[uuid("B196B288-BAB4-101A-B69C-00AA00341D07")]
    pub unsafe interface IOleControl: IUnknown {
        pub unsafe fn GetControlInfo(&self, param0: *mut OLECONTROLINFO) -> i32;
        pub unsafe fn OnMnemonic(&self, param0: *mut OLEACCELMSG) -> i32;
        pub unsafe fn OnAmbientPropertyChange(&self, param0: i32) -> HRESULT;
        pub unsafe fn FreezeEvents(&self, param0: i32) -> HRESULT;
    }

    #[uuid("B196B289-BAB4-101A-B69C-00AA00341D07")]
    pub unsafe interface IOleControlSite: IUnknown {
        pub unsafe fn OnControlInfoChanged(&self) -> HRESULT;
        pub unsafe fn LockInPlaceActive(&self, param0: i32) -> HRESULT;
        pub unsafe fn GetExtendedControl(&self, param0: *mut IDispatch) -> HRESULT;
        pub unsafe fn TransformCoords(&self, param0: i32, param1: i32, param2: i32) -> HRESULT;
        pub unsafe fn TranslateAccelerator(&self, param0: i32, param1: i32) -> i32;
        pub unsafe fn OnFocus(&self, param0: i32) -> HRESULT;
        pub unsafe fn ShowPropertyFrame(&self) -> HRESULT;
    }

    #[uuid("00000117-0000-0000-C000-000000000046")]
    pub unsafe interface IOleInPlaceActiveObject: IUnknownUnrestricted {
        pub unsafe fn GetWindow(&self, param0: *mut i32) -> i32;
        pub unsafe fn ContextSensitiveHelp(&self, param0: i32) -> i32;
        pub unsafe fn TranslateAccelerator(&self, param0: i32) -> i32;
        pub unsafe fn OnFrameWindowActivate(&self, param0: i32) -> i32;
        pub unsafe fn OnDocWindowActivate(&self, param0: i32) -> i32;
        pub unsafe fn ResizeBorder(&self, param0: i32, param1: *mut IOleInPlaceUIWindow, param2: i32) -> i32;
        pub unsafe fn EnableModeless(&self, param0: i32) -> i32;
    }

    #[uuid("00000115-0000-0000-C000-000000000046")]
    pub unsafe interface IOleInPlaceUIWindow: IOleWindow {
        pub unsafe fn GetBorder(&self, param0: i32) -> HRESULT;
        pub unsafe fn RequestBorderSpace(&self, param0: i32) -> HRESULT;
        pub unsafe fn SetBorderSpace(&self, param0: i32) -> HRESULT;
        pub unsafe fn SetActiveObject(&self, param0: *mut IOleInPlaceActiveObject, param1: *mut u16) -> HRESULT;
    }

    #[uuid("00000116-0000-0000-C000-000000000046")]
    pub unsafe interface IOleInPlaceFrame: IOleInPlaceUIWindow {
    }

    #[uuid("00000119-0000-0000-C000-000000000046")]
    pub unsafe interface IOleInPlaceSite: IOleWindow {
        pub unsafe fn CanInPlaceActivate(&self) -> i32;
        pub unsafe fn OnInPlaceActivate(&self) -> HRESULT;
        pub unsafe fn OnUIActivate(&self) -> HRESULT;
        pub unsafe fn GetWindowContext(&self, param0: *mut *mut IOleInPlaceFrame, param1: *mut *mut IOleInPlaceUIWindow, param2: i32, param3: i32, param4: i32) -> HRESULT;
        pub unsafe fn Scroll(&self, param0: CY) -> HRESULT;
        pub unsafe fn OnUIDeactivate(&self, param0: i32) -> HRESULT;
        pub unsafe fn OnInPlaceDeactivate(&self) -> HRESULT;
        pub unsafe fn DiscardUndoState(&self) -> HRESULT;
        pub unsafe fn DeactivateAndUndo(&self) -> HRESULT;
        pub unsafe fn OnPosRectChange(&self, param0: i32) -> HRESULT;
    }

    #[uuid("00020400-0000-0000-C000-000000000046")]
    pub unsafe interface IDispatch: IUnknown {
        pub unsafe fn GetTypeInfoCount(&self, param0: *mut i32) -> HRESULT;
        pub unsafe fn GetTypeInfo(&self, param0: i32, param1: i32, param2: *mut i32) -> HRESULT;
        pub unsafe fn GetIDsOfNames(&self, param0: *mut OLECLSID, param1: *mut *mut u16, param2: i32, param3: i32, param4: *mut i32) -> HRESULT;
        pub unsafe fn Invoke(&self, param0: i32, param1: *mut OLECLSID, param2: i32, param3: i16, param4: *mut DISPPARAMS, param5: *mut VARIANT, param6: *mut EXCEPINFO, param7: *mut i32) -> HRESULT;
    }

    #[uuid("376BD3AA-3845-101B-84ED-08002B2EC713")]
    pub unsafe interface IPerPropertyBrowsing: IUnknown {
        pub unsafe fn GetDisplayString(&self, param0: i32, param1: *mut i32) -> i32;
        pub unsafe fn MapPropertyToPage(&self, param0: i32, param1: *mut OLECLSID) -> i32;
        pub unsafe fn GetPredefinedStrings(&self, param0: i32, param1: *mut OLECALPOLESTR, param2: *mut OLECADWORD) -> i32;
        pub unsafe fn GetPredefinedValue(&self, param0: i32, param1: i32, param2: *mut VARIANT) -> i32;
    }

    #[uuid("B196B28B-BAB4-101A-B69C-00AA00341D07")]
    pub unsafe interface ISpecifyPropertyPages: IUnknown {
        pub unsafe fn GetPages(&self, param0: *mut OLECAUUID) -> HRESULT;
    }

    #[uuid("00020404-0000-0000-C000-000000000046")]
    pub unsafe interface IEnumVARIANTUnrestricted: IDispatch {
        pub unsafe fn Next(&self, param0: i32, param1: *mut VARIANT, param2: i32) -> HRESULT;
        pub unsafe fn Skip(&self, param0: i32) -> HRESULT;
        pub unsafe fn Reset(&self) -> HRESULT;
        pub unsafe fn Clone(&self, param0: *mut *mut IEnumVARIANT) -> HRESULT;
    }

    #[uuid("0000000B-0000-0000-C000-000000000046")]
    pub unsafe interface IStorage: IUnknown {
    }

    #[uuid("0000010E-0000-0000-C000-000000000046")]
    pub unsafe interface IDataObject: IUnknown {
    }

    #[uuid("CB5BDC81-93C1-11CF-8F20-00805F2CD064")]
    pub unsafe interface IObjectSafety: IUnknown {
        pub unsafe fn GetInterfaceSafetyOptions(&self, param0: *mut OLECLSID, param1: *mut i32, param2: *mut i32) -> HRESULT;
        pub unsafe fn SetInterfaceSafetyOptions(&self, param0: *mut OLECLSID, param1: i32, param2: i32) -> HRESULT;
    }

    #[uuid("38584260-0CFB-45E7-8FBB-5D20B311F5B8")]
    pub unsafe interface IOleInPlaceActiveObjectVB: IDispatch {
    }

    #[uuid("C895C8F9-6564-4123-8760-529F72AB9322")]
    pub unsafe interface IOleControlVB: IDispatch {
    }

    #[uuid("D5D3BBE3-DB60-4522-AF5B-D767FE736DDB")]
    pub unsafe interface IPerPropertyBrowsingVB: IDispatch {
    }

    #[uuid("00020D00-0000-0000-C000-000000000046")]
    pub unsafe interface IRichEditOle: IUnknown {
        pub unsafe fn GetClientSite(&self, param0: *mut *mut IOleClientSite) -> HRESULT;
        pub unsafe fn GetObjectCount(&self) -> i32;
        pub unsafe fn GetLinkCount(&self) -> i32;
        pub unsafe fn GetObject(&self, param0: i32, param1: *mut c_void, param2: i32) -> i32;
        pub unsafe fn InsertObject(&self, param0: *mut c_void) -> i32;
        pub unsafe fn ConvertObject(&self, param0: i32, param1: *mut OLECLSID, param2: *mut u8) -> i32;
        pub unsafe fn ActivateAs(&self, param0: *mut OLECLSID, param1: *mut OLECLSID) -> i32;
        pub unsafe fn SetHostNames(&self, param0: *mut u8, param1: *mut u8) -> i32;
        pub unsafe fn SetLinkAvailable(&self, param0: i32, param1: i32) -> i32;
        pub unsafe fn SetDvaspect(&self, param0: i32, param1: i32) -> i32;
        pub unsafe fn HandsOffStorage(&self, param0: i32) -> i32;
        pub unsafe fn SaveCompleted(&self, param0: i32, param1: *mut IStorage) -> i32;
        pub unsafe fn InPlaceDeactivate(&self) -> i32;
        pub unsafe fn ContextSensitiveHelp(&self, param0: i32) -> i32;
        pub unsafe fn GetClipboardData(&self, param0: i32, param1: i32, param2: *mut *mut IDataObject) -> i32;
        pub unsafe fn ImportDataObject(&self, param0: *mut IDataObject, param1: i16, param2: i32) -> i32;
    }

    #[uuid("00020D03-0000-0000-C000-000000000046")]
    pub unsafe interface IRichEditOleCallback: IUnknown {
        pub unsafe fn GetNewStorage(&self, param0: *mut *mut IStorage) -> HRESULT;
        pub unsafe fn GetInPlaceContext(&self, param0: *mut *mut IOleInPlaceFrame, param1: *mut *mut IOleInPlaceUIWindow, param2: *mut OLEINPLACEFRAMEINFO) -> HRESULT;
        pub unsafe fn ShowContainerUI(&self, param0: i32) -> HRESULT;
        pub unsafe fn QueryInsertObject(&self, param0: *mut OLECLSID, param1: *mut IStorage, param2: i32) -> HRESULT;
        pub unsafe fn DeleteObject(&self, param0: i32) -> HRESULT;
        pub unsafe fn QueryAcceptData(&self, param0: *mut IDataObject, param1: *mut i16, param2: i32, param3: i32, param4: i32) -> HRESULT;
        pub unsafe fn ContextSensitiveHelp(&self, param0: i32) -> HRESULT;
        pub unsafe fn GetClipboardData(&self, param0: i32, param1: i32, param2: *mut *mut IDataObject) -> HRESULT;
        pub unsafe fn GetDragDropEffect(&self, param0: i32, param1: i32, param2: *mut i32) -> HRESULT;
        pub unsafe fn GetContextMenu(&self, param0: i16, param1: i32, param2: i32, param3: *mut i32) -> HRESULT;
    }

}

impl IUnknownUnrestricted {}

impl IOleWindow {}

impl IOleClientSite {}

impl IOleObject {}

impl IOleInPlaceObject {}

impl IOleControl {}

impl IOleControlSite {}

impl IOleInPlaceActiveObject {}

impl IOleInPlaceUIWindow {}

impl IOleInPlaceFrame {}

impl IOleInPlaceSite {}

impl IDispatch {}

impl IPerPropertyBrowsing {}

impl ISpecifyPropertyPages {}

impl IEnumVARIANTUnrestricted {}

impl IStorage {}

impl IDataObject {}

impl IObjectSafety {}

impl IOleInPlaceActiveObjectVB {
    pub unsafe fn QueryInterface(
        &self,
        param0: *mut GUID,
        param1: *mut *mut c_void,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
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

    pub unsafe fn GetTypeInfoCount(&self, param0: *mut i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { plVal: param0 },
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
        param0: i32,
        param1: i32,
        param2: *mut i32,
    ) -> Result<(), HRESULT> {
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
                    Anonymous: VARIANT_0_0_0 { plVal: param2 },
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
        param0: *mut OLECLSID,
        param1: *mut *mut u16,
        param2: i32,
        param3: i32,
        param4: *mut i32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
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
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_LPWSTR.0 as u16,
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
                    Anonymous: VARIANT_0_0_0 { lVal: param2 },
                    ..Default::default()
                }),
            },
        });
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
                    Anonymous: VARIANT_0_0_0 { plVal: param4 },
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
        param1: *mut OLECLSID,
        param2: i32,
        param3: i16,
        param4: *mut DISPPARAMS,
        param5: *mut VARIANT,
        param6: *mut EXCEPINFO,
        param7: *mut i32,
    ) -> Result<(), HRESULT> {
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
                    Anonymous: VARIANT_0_0_0 { lVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I2.0 as u16,
                    Anonymous: VARIANT_0_0_0 { iVal: param3 },
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
                        byref: param6 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { plVal: param7 },
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

    pub unsafe fn TranslateAccelerator(
        &self,
        param0: *mut BOOL,
        param1: *mut i32,
        param2: i32,
        param3: i32,
        param4: i32,
        param5: i32,
        param6: i32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 { pboolVal: param0.0 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { plVal: param1 },
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
                    Anonymous: VARIANT_0_0_0 { lVal: param3 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param4 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param5 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param6 },
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
            0x1,
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

impl IOleControlVB {
    pub unsafe fn QueryInterface(
        &self,
        param0: *mut GUID,
        param1: *mut *mut c_void,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
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

    pub unsafe fn GetTypeInfoCount(&self, param0: *mut i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { plVal: param0 },
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
        param0: i32,
        param1: i32,
        param2: *mut i32,
    ) -> Result<(), HRESULT> {
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
                    Anonymous: VARIANT_0_0_0 { plVal: param2 },
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
        param0: *mut OLECLSID,
        param1: *mut *mut u16,
        param2: i32,
        param3: i32,
        param4: *mut i32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
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
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_LPWSTR.0 as u16,
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
                    Anonymous: VARIANT_0_0_0 { lVal: param2 },
                    ..Default::default()
                }),
            },
        });
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
                    Anonymous: VARIANT_0_0_0 { plVal: param4 },
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
        param1: *mut OLECLSID,
        param2: i32,
        param3: i16,
        param4: *mut DISPPARAMS,
        param5: *mut VARIANT,
        param6: *mut EXCEPINFO,
        param7: *mut i32,
    ) -> Result<(), HRESULT> {
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
                    Anonymous: VARIANT_0_0_0 { lVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I2.0 as u16,
                    Anonymous: VARIANT_0_0_0 { iVal: param3 },
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
                        byref: param6 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { plVal: param7 },
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

    pub unsafe fn GetControlInfo(
        &self,
        param0: *mut BOOL,
        param1: *mut i16,
        param2: *mut i32,
        param3: *mut i32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 { pboolVal: param0.0 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I2.0 as u16,
                    Anonymous: VARIANT_0_0_0 { piVal: param1 },
                    ..Default::default()
                }),
            },
        });
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
                    Anonymous: VARIANT_0_0_0 { plVal: param3 },
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
            0x1,
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

    pub unsafe fn OnMnemonic(
        &self,
        param0: *mut BOOL,
        param1: i32,
        param2: i32,
        param3: i32,
        param4: i32,
        param5: i32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 { pboolVal: param0.0 },
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
                    Anonymous: VARIANT_0_0_0 { lVal: param2 },
                    ..Default::default()
                }),
            },
        });
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
                    Anonymous: VARIANT_0_0_0 { lVal: param4 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { lVal: param5 },
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
            0x2,
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

impl IPerPropertyBrowsingVB {
    pub unsafe fn QueryInterface(
        &self,
        param0: *mut GUID,
        param1: *mut *mut c_void,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
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

    pub unsafe fn GetTypeInfoCount(&self, param0: *mut i32) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { plVal: param0 },
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
        param0: i32,
        param1: i32,
        param2: *mut i32,
    ) -> Result<(), HRESULT> {
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
                    Anonymous: VARIANT_0_0_0 { plVal: param2 },
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
        param0: *mut OLECLSID,
        param1: *mut *mut u16,
        param2: i32,
        param3: i32,
        param4: *mut i32,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
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
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_LPWSTR.0 as u16,
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
                    Anonymous: VARIANT_0_0_0 { lVal: param2 },
                    ..Default::default()
                }),
            },
        });
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
                    Anonymous: VARIANT_0_0_0 { plVal: param4 },
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
        param1: *mut OLECLSID,
        param2: i32,
        param3: i16,
        param4: *mut DISPPARAMS,
        param5: *mut VARIANT,
        param6: *mut EXCEPINFO,
        param7: *mut i32,
    ) -> Result<(), HRESULT> {
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
                    Anonymous: VARIANT_0_0_0 { lVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I2.0 as u16,
                    Anonymous: VARIANT_0_0_0 { iVal: param3 },
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
                        byref: param6 as *mut c_void,
                    },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_I4.0 as u16,
                    Anonymous: VARIANT_0_0_0 { plVal: param7 },
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

    pub unsafe fn GetDisplayString(
        &self,
        param0: *mut BOOL,
        param1: i32,
        param2: *mut BSTR,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 { pboolVal: param0.0 },
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
                    vt: ::windows::Win32::System::Ole::VT_BSTR.0 as u16,
                    Anonymous: VARIANT_0_0_0 {
                        pbstrVal: ::std::mem::transmute(param2),
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
            0x1,
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

    pub unsafe fn GetPredefinedStrings(
        &self,
        param0: *mut BOOL,
        param1: i32,
        param2: *mut *mut SAFEARRAY,
        param3: *mut *mut SAFEARRAY,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 { pboolVal: param0.0 },
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
        /* pointer to unknown type 0x1B */
        /* pointer to unknown type 0x1B */
        let mut disp_params = DISPPARAMS {
            rgvarg: arg_params.as_mut_ptr(),
            rgdispidNamedArgs: ::std::ptr::null_mut(),
            cArgs: arg_params.len() as u32,
            cNamedArgs: 0,
        };
        let invoke_result = IDispatch::Invoke(
            self,
            #[allow(overflowing_literals)]
            0x2,
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

    pub unsafe fn GetPredefinedValue(
        &self,
        param0: *mut BOOL,
        param1: i32,
        param2: i32,
        param3: *mut VARIANT,
    ) -> Result<(), HRESULT> {
        let mut arg_params = vec![];
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_BOOL.0 as u16,
                    Anonymous: VARIANT_0_0_0 { pboolVal: param0.0 },
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
                    Anonymous: VARIANT_0_0_0 { lVal: param2 },
                    ..Default::default()
                }),
            },
        });
        arg_params.push(VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: ::windows::Win32::System::Ole::VT_VARIANT.0 as u16,
                    Anonymous: VARIANT_0_0_0 { pvarVal: param3 },
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
            0x3,
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

impl IRichEditOle {}

impl IRichEditOleCallback {}
