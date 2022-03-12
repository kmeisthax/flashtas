//! Window Class Rust bridging
//!
//! This module contains a trait, `Window`, which establishes data ownership
//! rules for Win32 window classes.

use std::mem::forget;
use std::sync::{Arc, Mutex};
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{
    DefWindowProcW, GetWindowLongPtrW, SetWindowLongPtrW, CREATESTRUCTW, GWLP_USERDATA, WM_CREATE,
    WM_NCCREATE, WM_NCDESTROY,
};

/// A pointer to data associated with a Win32 HWND.
///
/// `Window`s have `Arc<Mutex>` semantics: the `Window` itself is a smart
/// pointer to an inner `Data` type. You may hold a `Window` to have guaranteed
/// access to it's data; while anyone with a still-valid `HWND` may also
/// recover access to the `Window`.
pub trait Window: Sized {
    /// The inner type of the window's data.
    type Data;

    /// Get the native window handle for this window.
    ///
    /// This function should panic if the `Window` has been destroyed.
    fn window(&self) -> HWND;

    /// Process a message from the window system.
    ///
    /// Incoming messages on the `Window`'s `HWND` will be processed and
    /// forwarded here. This should not be used to *send* window messages; you
    /// should use the standard Win32 messaging system for that.
    ///
    /// The function should return `None` for messages it does not handle.
    fn process_message(&self, msg: u32, _hparam: WPARAM, _lparam: LPARAM) -> Option<LRESULT>;

    /// Wrap a window data pointer as this impl of `Window`.
    fn from_data(data: Arc<Mutex<Self::Data>>) -> Self;

    /// Set the native window handle for this window.
    ///
    /// # Safety
    ///
    /// This must be a window handle obtained from a call to the window class's
    /// `WNDPROC` function, the handle must not already have a user data
    /// pointer, and the `Window` must have been recovered from the
    /// `lpCreateParams` of a `CREATESTRUCTW` of a `WM_CREATE` or `WM_NCCREATE`
    /// message. Said param should have come from a `CreateWindow` or
    /// `CreateWindowEx` which was populated with a raw pointer to an owned
    /// `Window::Data` reference.
    ///
    /// The `HWND` must have already had it's `USERDATA` pointer set.
    unsafe fn set_window_unchecked(&self, window: HWND);

    /// Forget the native window handle for this `Window`.
    fn forget_window(self);

    /// Create a `Window` from a native window handle for the first time.
    ///
    /// This parameter takes the `HWND` and `LPARAM` from the create message
    /// sent to the window class's `WNDPROC`.
    ///
    /// # Safety
    ///
    /// You may only call this function once per `Window` and only with the
    /// parameters of a `WM_CREATE` or `WM_NCCREATE` message sent to it. Those
    /// parameters must be valid (see `set_window` and
    /// `from_handle_unchecked`). You must not call `destroyed` before the
    /// window has been `created`.
    unsafe fn created(hwnd: HWND, lparam: LPARAM) -> Self {
        let create_struct_raw = lparam.0 as *mut CREATESTRUCTW;
        let dwindow_raw = (*create_struct_raw).lpCreateParams;

        SetWindowLongPtrW(hwnd, GWLP_USERDATA, dwindow_raw as isize);
        let wnd = Self::from_handle_unchecked(hwnd);

        wnd.set_window_unchecked(hwnd);

        wnd
    }

    /// Get the `Window` associated with a native window handle.
    ///
    /// # Safety
    ///
    /// The `HWND` must have been created by a `Window` whose `destroyed` method
    /// has not yet been called. The `Data` type of the `Window` must match the
    /// `Data` type owned by the `HWND`.
    unsafe fn from_handle_unchecked(window: HWND) -> Self {
        let data_ptr = GetWindowLongPtrW(window, GWLP_USERDATA) as *const Mutex<Self::Data>;
        let data_owned_by_hwnd = Arc::from_raw(data_ptr);

        // C pointers do not inherit Rust ownership semantics, so if we just
        // reconstitute the `Arc` then we can only call this function soundly
        // once. We instead need to clone the `Arc` so that the callee can
        // safely drop it, and then forget the HWND's `Arc` so that it doesn't
        // get dropped when we return.
        let data = data_owned_by_hwnd.clone();
        forget(data_owned_by_hwnd);

        Self::from_data(data)
    }

    /// Mark the `Window` as having been destroyed.
    ///
    /// This breaks the association between the `Window` and it's native
    /// `HWND`. It should only be called if the `HWND` will no longer be usable.
    ///
    /// # Safety
    ///
    /// You may only call this function once, and only after the `Window` has
    /// been `created`.
    ///
    /// You must not call `from_handle_unchecked` after calling this function.
    unsafe fn destroyed(self) {
        let window = self.window();

        let data_ptr = GetWindowLongPtrW(window, GWLP_USERDATA) as *const Mutex<Self::Data>;
        Arc::from_raw(data_ptr);

        SetWindowLongPtrW(window, GWLP_USERDATA, 0);
        self.forget_window();
    }

    /// The underlying window procedure for this `Window`'s class.
    extern "system" fn wndproc(hwnd: HWND, msg: u32, hparam: WPARAM, lparam: LPARAM) -> LRESULT {
        let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) };
        let wnd = if ptr == 0 {
            if msg == WM_CREATE || msg == WM_NCCREATE {
                unsafe { Self::created(hwnd, lparam) }
            } else {
                //TODO: How to handle WM_GETMINMAXINFO before HWND is fully initialized?
                return unsafe { DefWindowProcW(hwnd, msg, hparam, lparam) };
            }
        } else {
            unsafe { Self::from_handle_unchecked(hwnd) }
        };

        let param = match wnd.process_message(msg, hparam, lparam) {
            Some(p) => p,
            None => unsafe { DefWindowProcW(hwnd, msg, hparam, lparam) },
        };

        //TODO: This assumes no messages hit the window after destruction.
        if msg == WM_NCDESTROY {
            unsafe { wnd.destroyed() }
        }

        param
    }
}
