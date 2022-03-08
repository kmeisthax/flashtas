use crate::display::DisplayWindow;
use crate::unsafe_write_com;
use crate::window_class::Window;
use activex_rs::bindings::ole32::{
    IAdviseSink, IOleClientSite, IOleControlSite, IOleInPlaceActiveObject, IOleInPlaceFrame,
    IOleInPlaceSite, IOleInPlaceUIWindow, IOleWindow, BSTR, CY, OLEINPLACEFRAMEINFO,
};
use activex_rs::bindings::stdole::{IDispatch, IMoniker, IOleContainer};
use com::interfaces::IUnknown;
use com::production::ClassAllocation;
use lazy_static::lazy_static;
use std::mem::size_of;
use std::sync::{Arc, Mutex};
use windows::core::HRESULT;
use windows::Win32::Foundation::{HWND, RECT, S_OK};
use windows::Win32::System::Com::{CreateItemMoniker, FORMATETC, STGMEDIUM};
use windows::Win32::System::Ole::OleMenuGroupWidths;
use windows::Win32::UI::WindowsAndMessaging::{GetWindowRect, MSG};

com::interfaces! {
    /// COM interface for setting the display window.
    ///
    /// All the `Arc<Mutex>` juggling I did in `DisplayWindow` probably won't
    /// work with COM pointers so we're just going to have split objects.
    #[uuid("5210d570-8943-485a-b2c1-486730ba97cc")]
    pub unsafe interface ITASClientSite: IUnknown {
        pub fn attach_to_displaywnd(&self, new_wnd: *const DisplayWindow);
    }
}

com::class! {
    /// OLE client site for FlashTAS client controls
    pub class TASClientSite: ITASClientSite, IOleClientSite, IAdviseSink, IOleControlSite, IOleInPlaceFrame(IOleInPlaceUIWindow(IOleWindow)), IOleInPlaceSite(IOleWindow) {
        associated_display: Arc<Mutex<Option<DisplayWindow>>>
    }

    impl ITASClientSite for TASClientSite {
        fn attach_to_displaywnd(&self, new_wnd_raw: *const DisplayWindow) {
            let new_wnd = unsafe { &*new_wnd_raw };

            *self.associated_display.lock().unwrap() = Some(new_wnd.clone());
        }
    }

    impl IOleClientSite for TASClientSite {
        fn SaveObject(&self) -> HRESULT {
            unimplemented!();
        }

        fn GetMoniker(&self, param0: u32, param1: u32, param2: *mut IMoniker) -> HRESULT {
            let win_moniker = unsafe { CreateItemMoniker("!", "Flash") }.unwrap();

            if !param2.is_null() {
                unsafe_write_com!(param2, win_moniker);
            }

            S_OK
        }

        fn GetContainer(&self, param0: *mut IOleContainer) -> HRESULT {
            unimplemented!();
        }

        fn ShowObject(&self) -> HRESULT {
            S_OK
        }

        fn OnShowWindow(&self, param0: ::com::sys::BOOL) -> HRESULT {
            unimplemented!();
        }

        fn RequestNewObjectLayout(&self) -> HRESULT {
            unimplemented!();
        }
    }

    impl IAdviseSink for TASClientSite {
        pub unsafe fn OnDataChange(&self, param0: *mut FORMATETC, param1: *mut STGMEDIUM) {
            unimplemented!();
        }

        pub unsafe fn OnViewChange(&self, param0: u32, param1: i32) {
            unimplemented!();
        }

        pub unsafe fn OnRename(&self, param0: IMoniker) {
            unimplemented!();
        }

        pub unsafe fn OnSave(&self) {
            unimplemented!();
        }

        pub unsafe fn OnClose(&self) {
            unimplemented!();
        }
    }

    impl IOleControlSite for TASClientSite {
        pub unsafe fn OnControlInfoChanged(&self) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn LockInPlaceActive(&self, param0: i32) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn GetExtendedControl(&self, param0: IDispatch) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn TransformCoords(&self, param0: i32, param1: i32, param2: i32) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn TranslateAccelerator(&self, param0: i32, param1: i32) -> i32 {
            unimplemented!();
        }

        pub unsafe fn OnFocus(&self, param0: i32) -> HRESULT {
            S_OK
        }

        pub unsafe fn ShowPropertyFrame(&self) -> HRESULT {
            unimplemented!();
        }
    }

    impl IOleWindow for TASClientSite {
        pub unsafe fn GetWindow(&self, param0: *mut isize) -> HRESULT {
            let hwnd = self.associated_display.lock().unwrap().as_ref().unwrap().window();

            *param0 = hwnd.0;

            S_OK
        }

        pub unsafe fn ContextSensitiveHelp(&self, param0: i32) -> HRESULT {
            unimplemented!();
        }
    }

    impl IOleInPlaceUIWindow for TASClientSite {
        pub unsafe fn GetBorder(&self, param0: i32) -> HRESULT {
            unimplemented!();
        }
        pub unsafe fn RequestBorderSpace(&self, param0: i32) -> HRESULT {
            unimplemented!();
        }
        pub unsafe fn SetBorderSpace(&self, param0: i32) -> HRESULT {
            S_OK
        }
        pub unsafe fn SetActiveObject(&self, param0: IOleInPlaceActiveObject, param1: *mut u16) -> HRESULT {
            self.associated_display.lock().unwrap().as_ref().unwrap().set_active_object(param0);

            S_OK
        }
    }

    impl IOleInPlaceSite for TASClientSite {
        pub unsafe fn CanInPlaceActivate(&self) -> HRESULT {
            S_OK
        }

        pub unsafe fn OnInPlaceActivate(&self) -> HRESULT {
            S_OK
        }

        pub unsafe fn OnUIActivate(&self) -> HRESULT {
            S_OK
        }

        pub unsafe fn GetWindowContext(&self, param0: *mut IOleInPlaceFrame, param1: *mut IOleInPlaceUIWindow, param2: *mut RECT, param3: *mut RECT, param4: *mut OLEINPLACEFRAMEINFO) -> HRESULT {
            if !param0.is_null() {
                let ipf: IOleInPlaceFrame = self.into();
                unsafe_write_com!(param0, ipf);
            }

            if !param1.is_null() {
                let ipuiw: IOleInPlaceUIWindow = self.into();
                unsafe_write_com!(param1, ipuiw);
            }

            if !param2.is_null() || !param3.is_null() {
                let wnd = self.associated_display.lock().unwrap().as_ref().unwrap().window();
                let mut rect = RECT::default();
                GetWindowRect(wnd, &mut rect).unwrap();

                rect.bottom -= rect.top;
                rect.right -= rect.left;
                rect.top = 0;
                rect.left = 0;

                if !param2.is_null() {
                    *param2 = rect;
                }

                if !param3.is_null() {
                    *param3 = rect;
                }
            }

            if !param4.is_null() {
                let disp = self.associated_display.lock().unwrap();
                let disp_ref = disp.as_ref().unwrap();
                let ipfo = &mut *param4;

                //NOTE: Despite `repr(C)` we can't actually check if we have
                //space for individual fields because Rust doesn't support
                //pointer-to-subfields or struct offsets
                if ipfo.cb >= size_of::<OLEINPLACEFRAMEINFO>() {
                    ipfo.fMDIApp = ::windows::Win32::Foundation::BOOL::from(false);
                    ipfo.hWndFrame = disp_ref.window();

                    let (haccel, accel_count) = disp_ref.accel();

                    ipfo.hAccel = haccel;
                    ipfo.cAccelEntries = accel_count;
                }
            }

            S_OK
        }

        pub unsafe fn Scroll(&self, param0: CY) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn OnUIDeactivate(&self, param0: i32) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn OnInPlaceDeactivate(&self) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn DiscardUndoState(&self) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn DeactivateAndUndo(&self) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn OnPosRectChange(&self, param0: i32) -> HRESULT {
            unimplemented!();
        }
    }

    impl IOleInPlaceFrame for TASClientSite {
        pub unsafe fn InsertMenus(&self, param0: i64, param1: *mut OleMenuGroupWidths) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn SetMenu(&self, param0: i64, param1: i64, param2: i64) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn RemoveMenus(&self, param0: i64) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn SetStatusText(&self, param0: BSTR) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn EnableModeless(&self, param0: i32) -> HRESULT {
            unimplemented!();
        }

        pub unsafe fn TranslateAccelerator(&self, param0: *mut MSG, param1: u16) -> HRESULT {
            unimplemented!();
        }
    }
}

lazy_static! {
    pub static ref TASClientSite__CF: ClassAllocation<TASClientSiteClassFactory> =
        TASClientSiteClassFactory::allocate();
}
