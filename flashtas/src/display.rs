//! WNDCLASS for our display window.

use crate::tas_client::{ITASClientSite, TASClientSite__CF};
use crate::timer::Timer;
use crate::window_class::Window;
use activex_rs::bindings::flash::{
    IShockwaveFlash, IID___ISHOCKWAVE_FLASH_EVENTS, SHOCKWAVE_FLASH_CLSID,
};
use activex_rs::bindings::ole32::{
    IAdviseSink, IConnectionPointContainer, IOleClientSite, IOleInPlaceActiveObject,
    IOleInPlaceObject, IOleObject, IOleWindow,
};
use com::interfaces::IClassFactory;
use com::runtime::create_instance;
use flashtas_format::{AutomatedEvent, InputInjector, MouseButton, MouseButtons};
use lazy_static::lazy_static;
use std::ffi::c_void;
use std::ffi::OsStr;
use std::mem::MaybeUninit;
use std::os::windows::ffi::OsStrExt;
use std::path::Path;
use std::path::PathBuf;
use std::process::exit;
use std::ptr::{null, null_mut};
use std::sync::{Arc, Mutex};
use windows::core::Error as WinError;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, POINT, PWSTR, RECT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Ole::OLEIVERB_SHOW;
use windows::Win32::UI::Input::KeyboardAndMouse::VK_F4;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateAcceleratorTableW, CreateWindowExW, DispatchMessageW, GetClientRect, GetWindowRect,
    RegisterClassW, SetWindowPos, ACCEL, CW_USEDEFAULT, FALT, HACCEL, HMENU, MK_LBUTTON,
    MK_MBUTTON, MK_RBUTTON, MSG, SWP_NOMOVE, SWP_NOZORDER, WINDOW_EX_STYLE, WM_CREATE, WM_DESTROY,
    WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MBUTTONDOWN, WM_MBUTTONUP, WM_MOUSEMOVE, WM_RBUTTONDOWN,
    WM_RBUTTONUP, WM_SIZE, WM_USER, WNDCLASSW, WS_OVERLAPPEDWINDOW,
};

pub const DW_TIMER_ELAPSED: u32 = WM_USER;

/// A Window that holds an ActiveX/OLE control.
#[derive(Clone)]
pub struct DisplayWindow(Arc<Mutex<DisplayWindowData>>);

pub struct DisplayWindowData {
    window: HWND,

    movie: PathBuf,

    /// The current Flash instance.
    fp: Option<IShockwaveFlash>,

    /// The native window that corresponds to this Flash instance.
    fp_window: HWND,

    /// The movie's desired stage width.
    stage_width: i32,

    /// The movie's desired stage height.
    stage_height: i32,

    /// The input injector for this.
    injector: InputInjector,

    /// The current keyboard shortcut ("accelerator") table.
    accel: HACCEL,

    /// A timer object.
    timer: Timer,

    /// Number of times Flash Player has sent Readystate 4
    fp_ready: u8,

    /// Number of times Flash Player has gotten MSG 1025
    fp_got_message: bool,

    /// The last frame Flash Player reported.
    fp_lastframe: i32,
}

impl DisplayWindow {
    pub fn create<P: AsRef<Path>>(
        movie: PathBuf,
        width: i32,
        height: i32,
        frame_rate: f64,
        path: P,
    ) -> Result<Self, WinError> {
        let _ = *DISPLAY_WNDCLASS;

        let data = Arc::new(Mutex::new(DisplayWindowData {
            window: HWND(0),
            movie,
            fp: None,
            fp_window: HWND(0),
            stage_width: width,
            stage_height: height,
            injector: InputInjector::from_file(path).unwrap(),
            accel: HACCEL(0),
            timer: Timer::new(frame_rate),
            fp_ready: 0,
            fp_got_message: false,
            fp_lastframe: 0,
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
                width,
                height,
                HWND::default(),
                HMENU::default(),
                GetModuleHandleW(PWSTR(null())),
                Arc::into_raw(data.clone()) as *mut c_void,
            )
        }
        .ok()?;

        // If we put the window size in `CreateWindowExW` then the non-client
        // area gets subtracted from it which is wrong. So we have to actually
        // pull window metrics at this time and add it back.
        let mut client_rect = RECT::default();
        unsafe { GetClientRect(window, &mut client_rect).unwrap() };

        let mut nonclient_rect = RECT::default();
        unsafe { GetWindowRect(window, &mut nonclient_rect).unwrap() };

        let extra_width =
            (nonclient_rect.right - client_rect.right) + (client_rect.left - nonclient_rect.left);
        let extra_height =
            (nonclient_rect.bottom - client_rect.bottom) + (client_rect.top - nonclient_rect.top);

        unsafe {
            SetWindowPos(
                window,
                HWND(0),
                0,
                0,
                width + extra_width,
                height + extra_height,
                SWP_NOMOVE | SWP_NOZORDER,
            )
            .unwrap()
        };

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
        let mut child_wnd = 0;

        unsafe { object.GetWindow(&mut child_wnd).unwrap() };

        if child_wnd == 0 {
            eprintln!("Rejected attempt to set invalid child window");
            return;
        }

        self.0.lock().unwrap().fp_window = HWND(child_wnd);
    }

    /// Get the current active object's HWND.
    ///
    /// This will be 0 if the current active object has not been set yet.
    pub fn active_object_window(&self) -> HWND {
        self.0.lock().unwrap().fp_window
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

    /// Indicates that Flash Player has hit readystate 4.
    pub fn fp_ready(&self) {
        let could_start_pump = self.can_start_pump();

        self.0.lock().unwrap().fp_ready += 1;

        if !could_start_pump && self.can_start_pump() {
            self.start_pump();
        }
    }

    /// Indicates that sync messages (MSG 1025) has been sent to Flash Player.
    ///
    /// This also takes the HWND of the window that got this message. This
    /// allows intercepting Flash Player's messages earlier than when it says
    /// we're allowed to know what it's HWND is. This also relies on message
    /// 1025 being unique to FP.
    pub fn fp_got_message(&self, hwnd: HWND) {
        if !self.0.lock().unwrap().fp_got_message {
            self.0.lock().unwrap().fp_got_message = true;
            self.0.lock().unwrap().fp_window = hwnd;

            // For some reason, we *always* need to pump one frame on MSG 1025,
            // even if it would otherwise mess with our sync, otherwise we can't
            // hit frame 2.
            log::debug!(
                "FP got first message, pumping frame {} IMMEDIATELY",
                self.frame_ctr()
            );
            self.pump();
        }
    }

    fn can_start_pump(&self) -> bool {
        let self_mut = self.0.lock().unwrap();

        self_mut.fp_ready >= 2 || self_mut.fp_got_message
    }

    /// Start the timer that drives synthetic event pumping.
    fn start_pump(&self) {
        let window = self.window();

        let self_mut = self.0.lock().unwrap();
        unsafe {
            self_mut.timer.set_message(window, DW_TIMER_ELAPSED);
        }
    }

    /// Pump the next frame's worth of synthetic inputs into the player.
    pub fn pump(&self) {
        let mut hwnd = self.active_object_window();
        if hwnd.is_invalid() {
            // At this point we absolutely should have a Flash Player window, even
            // if Flash Player hasn't told us about it yet.
            let fp = self.0.lock().unwrap().fp.clone().unwrap();
            let fp_ole = fp.query_interface::<IOleWindow>().unwrap();
            let mut fp_window = 0;

            unsafe { fp_ole.GetWindow(&mut fp_window).unwrap() };

            if fp_window != 0 {
                self.0.lock().unwrap().fp_window = HWND(fp_window);
                hwnd = HWND(fp_window);
            } else {
                log::error!("Cannot pump events into windowless player");
                return;
            }
        }

        let mut client_rect = RECT::default();
        unsafe { GetClientRect(hwnd, &mut client_rect) }.unwrap();

        let (stage_width, stage_height) = {
            let me = self.0.lock().unwrap();
            (me.stage_width, me.stage_height)
        };

        let stage_aspect_ratio = stage_width as f64 / stage_height as f64;
        let client_aspect_ratio = client_rect.right as f64 / client_rect.bottom as f64;

        let (offset_x, offset_y, scale) = if client_aspect_ratio > stage_aspect_ratio {
            let scale = stage_height as f64 / client_rect.bottom as f64;
            let offset_x = (client_rect.right as f64 - (stage_width as f64 * scale)) / 2.0;

            (offset_x as i16, 0, scale)
        } else {
            let scale = stage_width as f64 / client_rect.right as f64;
            let offset_y = (client_rect.bottom as f64 - (stage_height as f64 * scale)) / 2.0;

            (0, offset_y as i16, scale)
        };

        let mut window_rect = RECT::default();
        unsafe { GetWindowRect(hwnd, &mut window_rect) }.unwrap();

        // We cannot send messages to Flash Player while it's parent window is
        // locked by Rust.
        let mut events = vec![];

        self.0.lock().unwrap().injector.next(|evt, buttons| {
            let mut buttons_wparam = 0;
            if buttons.contains(MouseButtons::LEFT) {
                buttons_wparam |= MK_LBUTTON;
            }

            if buttons.contains(MouseButtons::MIDDLE) {
                buttons_wparam |= MK_MBUTTON;
            }

            if buttons.contains(MouseButtons::RIGHT) {
                buttons_wparam |= MK_RBUTTON;
            }

            let (pos, message) = match evt {
                AutomatedEvent::MouseMove { pos } => (pos, WM_MOUSEMOVE),
                AutomatedEvent::MouseDown { pos, btn } => match btn {
                    MouseButton::Left => (pos, WM_LBUTTONDOWN),
                    MouseButton::Middle => (pos, WM_MBUTTONDOWN),
                    MouseButton::Right => (pos, WM_RBUTTONDOWN),
                },
                AutomatedEvent::MouseUp { pos, btn } => match btn {
                    MouseButton::Left => (pos, WM_LBUTTONUP),
                    MouseButton::Middle => (pos, WM_MBUTTONUP),
                    MouseButton::Right => (pos, WM_RBUTTONUP),
                },
                AutomatedEvent::Wait => unreachable!(),
            };

            let client_x = offset_x + (pos.0 as f64 * scale) as i16;
            let client_y = offset_y + (pos.1 as f64 * scale) as i16;
            let pos_lparam = (client_x as u16 as u32 | (client_y as u16 as u32) << 16) as isize;

            //TODO: This assumes the child window has no non-client area.
            //Is that always true?
            let pt = POINT {
                x: window_rect.left + client_x as i32,
                y: window_rect.top + client_y as i32,
            };

            events.push(MSG {
                hwnd,
                message,
                wParam: WPARAM(buttons_wparam as usize),
                lParam: LPARAM(pos_lparam),
                time: 0, // TODO: timestamping
                pt,
            });
        });

        for msg in events {
            unsafe { DispatchMessageW(&msg) };
        }
    }

    /// Get the current Flash Player timeline frame.
    pub fn frame_ctr(&self) -> i32 {
        let self_mut = self.0.lock().unwrap();
        let fp = self_mut.fp.as_ref().unwrap();

        unsafe { fp.CurrentFrame() }.unwrap()
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

    fn process_message(&self, msg: u32, _hparam: WPARAM, lparam: LPARAM) -> Option<LRESULT> {
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
                log::info!("Loading SWF: {:?}", movie.as_os_str());

                let fp_conn = fp.query_interface::<IConnectionPointContainer>().unwrap();

                unsafe {
                    let mut fp_conn_pt = MaybeUninit::uninit();

                    fp_conn
                        .FindConnectionPoint(IID___ISHOCKWAVE_FLASH_EVENTS, fp_conn_pt.as_mut_ptr())
                        .unwrap();

                    //TODO: How do you check for nullptr in MaybeUninit without UB?
                    let fp_conn_pt = fp_conn_pt.assume_init();
                    let mut cookie = 0;
                    fp_conn_pt.Advise(client_site.clone(), &mut cookie).unwrap();

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
                // No, it's not because the client site wasn't being retained.
                //
                // THIS IS NOT THE CORRECT WAY TO DO THINGS IN RUST.
                ::std::mem::forget(fp_ole);
                ::std::mem::forget(fp);

                Some(LRESULT(0))
            }
            WM_SIZE => {
                let fp = self.0.lock().unwrap().fp.clone().unwrap();
                let fp_ole_inplace = fp.query_interface::<IOleInPlaceObject>().unwrap();

                let new_width = lparam.0 & 0xFFFF;
                let new_height = (lparam.0 >> 16) & 0xFFFF;

                let rect = RECT {
                    top: 0,
                    left: 0,
                    right: new_width as i32,
                    bottom: new_height as i32,
                };

                unsafe { fp_ole_inplace.SetObjectRects(&rect, &rect).unwrap() };

                Some(LRESULT(0))
            }
            WM_DESTROY => exit(0),
            DW_TIMER_ELAPSED => {
                let cur_frame = self.frame_ctr();
                let last_frame = self.0.lock().unwrap().fp_lastframe;

                log::debug!("Timer fired, last frame was {}", last_frame);

                if cur_frame != last_frame {
                    log::debug!("Frame updated to {}, pumping new frame", cur_frame);
                    self.pump();
                }

                self.0.lock().unwrap().fp_lastframe = cur_frame;

                Some(LRESULT(0))
            }
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
