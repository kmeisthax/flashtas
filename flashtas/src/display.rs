//! WNDCLASS for our display window.

use crate::tas_client::{ITASClientSite, TASClientSite, TASClientSite__CF};
use activex_rs::bindings::flash::{IShockwaveFlash, SHOCKWAVE_FLASH_CLSID};
use activex_rs::bindings::ole32::{
    IAdviseSink, IOleClientSite, IOleInPlaceActiveObject, IOleObject, IOleWindow, IViewObject,
};
use com::interfaces::IClassFactory;
use com::runtime::create_instance;
use lazy_static::lazy_static;
use std::ffi::c_void;
use std::ffi::OsStr;
use std::mem::forget;
use std::os::windows::ffi::OsStrExt;
use std::process::exit;
use std::ptr::{null, null_mut};
use std::sync::{Arc, Mutex};
use windows::core::Error as WinError;
use windows::Win32::Foundation::{BOOL, HWND, LPARAM, LRESULT, PWSTR, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Ole::OLEIVERB_SHOW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, GetWindowLongPtrW, RegisterClassW, SetWindowLongPtrW,
    CREATESTRUCTW, CW_USEDEFAULT, GWLP_USERDATA, HMENU, WINDOW_EX_STYLE, WM_CREATE, WM_DESTROY,
    WM_NCCREATE, WM_NCDESTROY, WNDCLASSW, WS_OVERLAPPEDWINDOW,
};

/// The window class of our main display window.
#[derive(Clone)]
pub struct DisplayWindow(Arc<Mutex<DisplayWindowData>>);

pub struct DisplayWindowData {
    window: HWND,

    /// The current Flash instance.
    fp: Option<IShockwaveFlash>,
}

impl DisplayWindow {
    pub fn create() -> Result<Self, WinError> {
        let _ = *DISPLAY_WNDCLASS;

        let data = Arc::new(Mutex::new(DisplayWindowData {
            window: HWND(0),
            fp: None,
        }));

        // The HWND itself owns an `Arc<Mutex<Self>>` through the C pointer,
        // which we send via `lpParam`. The `WNDPROC` will grab the param from
        // the first message it gets
        let window = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                PWSTR(DISPLAY_CLASSNAME.as_ptr()),
                PWSTR(DISPLAY_WINNAME.as_ptr()),
                WS_OVERLAPPEDWINDOW,
                CW_USEDEFAULT,
                CW_USEDEFAULT,
                CW_USEDEFAULT,
                CW_USEDEFAULT,
                HWND::default(),
                HMENU::default(),
                GetModuleHandleW(PWSTR(null())),
                Arc::into_raw(data.clone()) as *mut c_void,
            )
        }
        .ok()?;

        // Wrap the HWND back into the data.
        //
        // We specifically do it in this order because `CreateWindowExW` can
        // send messages to ourselves *before* it returns, so it's better for
        // the WNDPROC to see a valid `DisplayWindow` with a missing `HWND`
        // than a dangling pointer.
        data.lock().unwrap().window = window;

        Ok(Self(data))
    }

    /// Retrieve this `DisplayWindow`'s `HWND`.
    pub fn window(&self) -> HWND {
        self.0.lock().unwrap().window
    }

    /// Retrieve the `DisplayWindow` from an `HWND`.
    ///
    /// # Safety
    ///
    /// The `HWND` must originate from a `DisplayWindow` whose `destroy` method
    /// has not yet been called.
    pub unsafe fn from_handle_unchecked(window: HWND) -> Self {
        let data_ptr = GetWindowLongPtrW(window, GWLP_USERDATA) as *const Mutex<DisplayWindowData>;
        let data_owned_by_hwnd = Arc::from_raw(data_ptr);

        // C pointers do not inherit Rust ownership semantics, so if we just
        // reconstitute the `Arc` then we can only call this function soundly
        // once. We instead need to clone the `Arc` so that the callee can
        // safely drop it, and then forget the HWND's `Arc` so that it doesn't
        // get dropped when we return.
        let data = data_owned_by_hwnd.clone();
        forget(data_owned_by_hwnd);

        Self(data)
    }

    /// Mark this `DisplayWindow` as destroyed by Windows to avoid a memory
    /// leak.
    ///
    /// Calling this function breaks the association between a `DisplayWindow`
    /// and it's `HWND`. The `DisplayWindow`'s `window` field will be nulled
    /// out, and the `HWND`'s internal reference to the `DisplayWindow` will be
    /// dropped.
    ///
    /// # Safety
    ///
    /// The `HWND` must originate from a `DisplayWindow`. You must not call
    /// `destroy` more than once.
    pub unsafe fn destroy(self) {
        let window = self.window();

        let data_ptr = GetWindowLongPtrW(window, GWLP_USERDATA) as *const Mutex<DisplayWindowData>;
        Arc::from_raw(data_ptr);

        SetWindowLongPtrW(window, GWLP_USERDATA, 0);
        self.0.lock().unwrap().window = HWND(0);
    }

    /// Process a message from the window system.
    ///
    /// This should not be used to *send* window messages; you should use the
    /// standard Win32 messaging system for that.
    ///
    /// The function should return `None` for messages it does not handle.
    pub fn process_message(&self, msg: u32, _hparam: WPARAM, _lparam: LPARAM) -> Option<LRESULT> {
        match msg {
            WM_CREATE => {
                let fp = create_instance::<IShockwaveFlash>(&SHOCKWAVE_FLASH_CLSID).expect("Flash");

                println!(
                    "Flash version: 0x{:x}",
                    unsafe { fp.FlashVersion() }.expect("Flash version")
                );

                let fp_ole = fp
                    .query_interface::<IOleObject>()
                    .expect("Must be an OLE object too");

                let class_factory = IClassFactory::from(&***TASClientSite__CF);
                let tas_client_site = class_factory.create_instance::<ITASClientSite>().unwrap();
                unsafe { tas_client_site.attach_to_displaywnd(self) };

                let client_site = tas_client_site.query_interface::<IOleClientSite>().unwrap();
                let advise_sink = tas_client_site.query_interface::<IAdviseSink>().unwrap();

                unsafe {
                    fp_ole.SetClientSite(client_site.clone()).unwrap();
                    fp_ole.Advise(advise_sink, null_mut()).unwrap();
                    fp_ole
                        .DoVerb(OLEIVERB_SHOW, null_mut(), client_site, 0, 0, null())
                        .unwrap();
                    fp.LoadMovie(
                        0,
                        OsStr::new("test.swf\0")
                            .encode_wide()
                            .collect::<Vec<_>>()
                            .leak()
                            .as_ptr(),
                    )
                    .unwrap();
                }

                self.0.lock().unwrap().fp = Some(fp);

                Some(LRESULT(0))
            }
            WM_DESTROY => exit(0),
            _ => None,
        }
    }
}

lazy_static! {
    static ref DISPLAY_WNDCLASS: u16 = unsafe {
        RegisterClassW(&WNDCLASSW {
            lpfnWndProc: Some(wndproc),
            hInstance: GetModuleHandleW(PWSTR(null())),
            lpszClassName: PWSTR(DISPLAY_CLASSNAME.as_ptr()),
            ..Default::default()
        })
    };
    static ref DISPLAY_CLASSNAME: &'static [u16] = OsStr::new("FlashTASDisplay\0")
        .encode_wide()
        .collect::<Vec<_>>()
        .leak();
    static ref DISPLAY_WINNAME: &'static [u16] = OsStr::new("FlashTAS\0")
        .encode_wide()
        .collect::<Vec<_>>()
        .leak();
}

extern "system" fn wndproc(hwnd: HWND, msg: u32, hparam: WPARAM, lparam: LPARAM) -> LRESULT {
    let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) };
    if ptr == 0 && (msg == WM_CREATE || msg == WM_NCCREATE) {
        let create_struct_raw = lparam.0 as *mut CREATESTRUCTW;
        let dwindow_raw = unsafe { &mut *create_struct_raw }.lpCreateParams;

        unsafe { SetWindowLongPtrW(hwnd, GWLP_USERDATA, dwindow_raw as isize) };
    } else if ptr == 0 {
        //TODO: How to handle WM_GETMINMAXINFO before HWND is fully initialized?
        return unsafe { DefWindowProcW(hwnd, msg, hparam, lparam) };
    }

    let wnd = unsafe { DisplayWindow::from_handle_unchecked(hwnd) };

    let param = match wnd.process_message(msg, hparam, lparam) {
        Some(p) => p,
        None => unsafe { DefWindowProcW(hwnd, msg, hparam, lparam) },
    };

    //TODO: This assumes no messages hit the window after destruction.
    if msg == WM_NCDESTROY {
        unsafe { wnd.destroy() }
    }

    param
}
