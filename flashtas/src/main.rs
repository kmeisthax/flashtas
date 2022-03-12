use crate::window_class::Window;
use clap::Parser;
use std::fs::{canonicalize, File};
use std::mem::size_of;
use std::path::{Component, PathBuf, Prefix};
use std::ptr::null_mut;
use std::time::Instant;
use swf::decompress_swf;
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Ole::OleInitialize;
use windows::Win32::System::Threading::{GetStartupInfoW, STARTUPINFOW};
use windows::Win32::UI::WindowsAndMessaging::{
    DispatchMessageW, GetMessageW, ShowWindow, TranslateMessage, MSG, SHOW_WINDOW_CMD, SW_HIDE,
    SW_SHOWNORMAL, WM_LBUTTONDBLCLK, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MBUTTONDBLCLK,
    WM_MBUTTONDOWN, WM_MBUTTONUP, WM_MOUSEHWHEEL, WM_MOUSEMOVE, WM_MOUSEWHEEL, WM_RBUTTONDBLCLK,
    WM_RBUTTONDOWN, WM_RBUTTONUP, WM_XBUTTONDBLCLK, WM_XBUTTONDOWN, WM_XBUTTONUP,
};

mod display;
mod tas_client;
mod timer;
mod window_class;

#[macro_use]
mod glue;

#[derive(Parser)]
struct Args {
    /// Path to the movie to run.
    movie: String,

    /// Path to the input file to hit.
    input: String,
}

fn main() {
    unsafe { OleInitialize(null_mut()) }.expect("Initialized OLE session");

    env_logger::init();

    let args = Args::parse();

    // Flash Player wants canonicalized paths, but chokes on what Rust calls
    // `VerbatimDisk` (e.g. \\?\C:\test.swf), so we have to *de*canonicalize it
    // slightly for FP.
    let movie_canon = canonicalize(args.movie).unwrap();
    let mut movie = PathBuf::new();
    for component in movie_canon.components() {
        movie.push(match component {
            Component::Prefix(p) => match p.kind() {
                Prefix::VerbatimDisk(d) => {
                    movie.push(format!("{}:/", ::std::char::from_u32(d as u32).unwrap()));
                    continue;
                }
                _ => Component::Prefix(p),
            },
            c => c,
        });
    }

    // Getting the stage size from Flash Player is too hard. Let's just parse
    // the SWF ourselves!
    let swf_header = decompress_swf(File::open(movie_canon).unwrap()).unwrap();
    let stage_size = swf_header.header.stage_size();
    let width = stage_size.x_max.to_pixels() as i32;
    let height = stage_size.y_max.to_pixels() as i32;
    let frame_rate = swf_header.header.frame_rate().to_f64();
    log::info!("Desired stage dimensions: {}x{}", width, height);

    let mainwnd =
        display::DisplayWindow::create(movie, width, height, frame_rate, args.input).unwrap();
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

    let mut last;

    while unsafe { GetMessageW(&mut msg, HWND::default(), 0, 0) }.as_bool() {
        last = Instant::now();

        unsafe { TranslateMessage(&msg) };

        // Filter all queued Flash Player messages, we'll be adding our own.
        if !msg.hwnd.is_invalid() && msg.hwnd == mainwnd.active_object_window() {
            match msg.message {
                WM_LBUTTONDOWN | WM_LBUTTONUP | WM_LBUTTONDBLCLK | WM_MBUTTONDBLCLK
                | WM_MBUTTONDOWN | WM_MBUTTONUP | WM_MOUSEHWHEEL | WM_MOUSEMOVE | WM_MOUSEWHEEL
                | WM_RBUTTONDBLCLK | WM_RBUTTONDOWN | WM_RBUTTONUP | WM_XBUTTONDBLCLK
                | WM_XBUTTONDOWN | WM_XBUTTONUP => continue,
                _ => {}
            }
        } else if mainwnd.active_object_window().is_invalid() && msg.message == 1025 {
            //Assume message 1025 ONLY gets sent to FP.
            mainwnd.fp_got_message(msg.hwnd);
        }

        log::debug!("Message: {} to {:?}", msg.message, msg.hwnd);
        log::debug!("Flash Player is {:?}", mainwnd.active_object_window());

        unsafe { DispatchMessageW(&msg) };

        let next = Instant::now();
        log::debug!("Process time: {:?}", next - last);
    }
}
