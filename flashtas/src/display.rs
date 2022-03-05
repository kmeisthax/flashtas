//! WNDCLASS for our display window.

use crate::tas_client::{ITASClientSite, TASClientSite__CF};
use crate::window_class::Window;
use activex_rs::bindings::flash::{IShockwaveFlash, SHOCKWAVE_FLASH_CLSID};
use activex_rs::bindings::ole32::{IAdviseSink, IOleClientSite, IOleObject};
use com::interfaces::IClassFactory;
use com::runtime::create_instance;
use lazy_static::lazy_static;
use std::ffi::c_void;
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use std::process::exit;
use std::ptr::{null, null_mut};
use std::sync::{Arc, Mutex};
use windows::core::Error as WinError;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, PWSTR, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Ole::OLEIVERB_SHOW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, RegisterClassW, CW_USEDEFAULT, HMENU, WINDOW_EX_STYLE, WM_CREATE, WM_DESTROY,
    WNDCLASSW, WS_OVERLAPPEDWINDOW,
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
}

impl Window for DisplayWindow {
    type Data = DisplayWindowData;

    fn window(&self) -> HWND {
        let wnd = self.0.lock().unwrap().window;
        if wnd.is_invalid() {
            panic!("Window already destroyed");
        }

        wnd
    }

    fn from_data(data: Arc<Mutex<Self::Data>>) -> Self {
        Self(data)
    }

    unsafe fn set_window_unchecked(&self, window: HWND) {
        self.0.lock().unwrap().window = window;
    }

    fn forget_window(self) {
        self.0.lock().unwrap().window = HWND(0);
    }

    fn process_message(&self, msg: u32, _hparam: WPARAM, _lparam: LPARAM) -> Option<LRESULT> {
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
            lpfnWndProc: Some(DisplayWindow::wndproc),
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
