//! WNDCLASS for our display window.

use crate::tas_client::{ITASClientSite, TASClientSite__CF};
use crate::window_class::Window;
use activex_rs::bindings::flash::{IShockwaveFlash, SHOCKWAVE_FLASH_CLSID};
use activex_rs::bindings::ole32::{
    IAdviseSink, IOleClientSite, IOleInPlaceActiveObject, IOleObject,
};
use com::interfaces::IClassFactory;
use com::runtime::create_instance;
use lazy_static::lazy_static;
use std::ffi::c_void;
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use std::path::PathBuf;
use std::process::exit;
use std::ptr::{null, null_mut};
use std::sync::{Arc, Mutex};
use windows::core::Error as WinError;
use windows::Win32::Foundation::{BOOL, HWND, LPARAM, LRESULT, PWSTR, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Ole::OLEIVERB_SHOW;
use windows::Win32::UI::Input::KeyboardAndMouse::VK_F4;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateAcceleratorTableW, CreateWindowExW, RegisterClassW, ACCEL, CW_USEDEFAULT, FALT, HACCEL,
    HMENU, WINDOW_EX_STYLE, WM_CREATE, WM_DESTROY, WNDCLASSW, WS_OVERLAPPEDWINDOW,
};

/// A Window that holds an ActiveX/OLE control.
#[derive(Clone)]
pub struct DisplayWindow(Arc<Mutex<DisplayWindowData>>);

pub struct DisplayWindowData {
    window: HWND,

    movie: PathBuf,

    /// The current Flash instance.
    fp: Option<IShockwaveFlash>,

    /// The current keyboard shortcut ("accelerator") table.
    accel: HACCEL,
}

impl DisplayWindow {
    pub fn create(movie: PathBuf) -> Result<Self, WinError> {
        let _ = *DISPLAY_WNDCLASS;

        let data = Arc::new(Mutex::new(DisplayWindowData {
            window: HWND(0),
            movie,
            fp: None,
            accel: HACCEL(0),
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

    /// Set the active object for this display window.
    pub fn set_active_object(&self, object: IOleInPlaceActiveObject) {
        let child_wnd = 0;
        unsafe { object.OnFrameWindowActivate(BOOL::from(true).0).unwrap() };
    }

    pub fn accel(&self) -> (HACCEL, usize) {
        let mut me = self.0.lock().unwrap();

        if me.accel.is_invalid() {
            let accel_table = vec![ACCEL {
                fVirt: FALT as u8,
                key: VK_F4.0,
                cmd: 0,
            }];

            me.accel =
                unsafe { CreateAcceleratorTableW(accel_table.as_ptr(), accel_table.len() as i32) };
        }

        //TODO: Actually get the accel table length
        (me.accel, 1)
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
                self.0.lock().unwrap().fp = Some(fp.clone());

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

                let movie = self.0.lock().unwrap().movie.clone();
                eprintln!("Loading SWF: {:?}", movie.as_os_str());

                unsafe {
                    fp_ole.SetClientSite(client_site.clone()).unwrap();
                    fp_ole.Advise(advise_sink, null_mut()).unwrap();
                    fp_ole
                        .DoVerb(
                            OLEIVERB_SHOW,
                            null_mut(),
                            client_site,
                            0,
                            self.window().0 as i64,
                            null(),
                        )
                        .unwrap();
                    fp.LoadMovie(
                        0,
                        movie
                            .as_os_str()
                            .encode_wide()
                            .collect::<Vec<_>>()
                            .leak()
                            .as_ptr(),
                    )
                    .unwrap();
                }

                // EVIL: For some reason, the COM pointers we currently hold
                // are somehow unbalanced, and we can't release them normally.
                // Don't ask me why, just never call `Release` on an Flash.ocx
                // class.
                //
                // THIS IS NOT THE CORRECT WAY TO DO THINGS IN RUST.
                ::std::mem::forget(fp_ole);
                ::std::mem::forget(fp);

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
