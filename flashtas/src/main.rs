use std::mem::size_of;
use std::ptr::null_mut;
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Ole::OleInitialize;
use windows::Win32::System::Threading::{GetStartupInfoW, STARTUPINFOW};
use windows::Win32::UI::WindowsAndMessaging::{
    DispatchMessageW, GetMessageW, ShowWindow, TranslateMessage, MSG, SHOW_WINDOW_CMD, SW_HIDE,
    SW_SHOWNORMAL,
};

mod display;
mod tas_client;

fn main() {
    unsafe { OleInitialize(null_mut()) }.expect("Initialized OLE session");

    let mainwnd = display::DisplayWindow::create().unwrap();
    let mut si = STARTUPINFOW {
        cb: size_of::<STARTUPINFOW>() as u32,
        ..Default::default()
    };

    unsafe {
        GetStartupInfoW(&mut si);
    }

    // For some reason the VSCode terminal sets `nCmdShow` to `SW_HIDE`, which
    // means windows never actually display. :/
    if si.wShowWindow == SW_HIDE.0 as u16 {
        si.wShowWindow = SW_SHOWNORMAL.0 as u16;
    }

    unsafe {
        ShowWindow(mainwnd.window(), SHOW_WINDOW_CMD(si.wShowWindow as u32));
    }

    let mut msg = MSG::default();

    while unsafe { GetMessageW(&mut msg, HWND::default(), 0, 0) }.as_bool() {
        unsafe {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
}
