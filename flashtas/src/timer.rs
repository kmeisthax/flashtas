//! Windows timers
//!
//! See the documentation for `Timer` for more information.

use std::ffi::c_void;
use std::mem::forget;
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};
use windows::Win32::Foundation::{BOOLEAN, HANDLE, HWND, LPARAM, WPARAM};
use windows::Win32::Media::timeBeginPeriod;
use windows::Win32::System::Threading::{
    CreateTimerQueueTimer, DeleteTimerQueueTimer, WORKER_THREAD_FLAGS,
};
use windows::Win32::UI::WindowsAndMessaging::PostMessageW;

/// A thin wrapper over Windows's timer queues.
///
/// # Drift correction
///
/// Features drift correction to deal with OSes that insist on their timer APIs
/// not accounting for it.
///
/// Our drift correction strategy is to:
///
///  - Run the timer queue at twice the desired frequency
///  - Measure the time between each timer firing, using Rust's `Instant` and
///    `Duration` types (which are themselves wrappers over equivalent Windows
///    APIs) and report the timer firing only when the elapsed period has been
///    met.
///
/// If the system timer is honest, we will effectively divide the frequency by
/// two and get the correct timing. If it's dishonest, we will trade off drift
/// for jitter by running timers early to avoid being late.
///
/// # Threading & messaging
///
/// For whatever reason, Windows timers don't fire on the message queue, so we
/// require that the caller provide their HWND and a user message ID to send.
/// This message will then be sent whenever the underlying timer is fired. You
/// must then call `Timer.elapsed` to actually determine if the period has
/// elapsed or not.
#[derive(Clone)]
pub struct Timer(Arc<Mutex<TimerData>>);

struct TimerData {
    /// The native timer handle.
    timer: HANDLE,

    /// The last time the timer was fired.
    ///
    /// If the timer has not yet been fired, this is the time it was created.
    last: Instant,

    /// How much additional time has elapsed beyond the last time the timer was
    /// fired.
    drift: Duration,

    /// The target period for the timer.
    target_period: Duration,

    /// The HWND of the caller to send messages to.
    caller_hwnd: HWND,

    /// A reserved user message used to flag the caller that it's timer may
    /// have elapsed.
    caller_msg: u32,
}

impl Timer {
    /// Create a new timer for a given frame rate.
    ///
    /// Make sure to call `set_message` in order to actually receive timer
    /// messages.
    pub fn new(frame_rate_target: f64) -> Self {
        let target_period = Duration::from_secs_f64(1.0 / frame_rate_target);
        let me = Arc::new(Mutex::new(TimerData {
            timer: HANDLE(0),
            last: Instant::now(),
            drift: Duration::new(0, 0),
            target_period,
            caller_hwnd: HWND(0),
            caller_msg: 0,
        }));

        let mut timer = HANDLE(0);
        let period = (target_period / 2).as_millis();

        unsafe {
            timeBeginPeriod(period as u32);
            CreateTimerQueueTimer(
                &mut timer,
                HANDLE(0),
                Some(Self::timer_callback),
                Arc::into_raw(me.clone()) as *mut c_void,
                0,
                period as u32,
                WORKER_THREAD_FLAGS(0),
            );
        }

        me.lock().unwrap().timer = timer;

        Self(me)
    }

    /// Set the owner's HWND and message.
    ///
    /// This also resets the timer's internal timing information, as if it had
    /// been created immediately.
    ///
    /// # Safety
    ///
    /// `caller_hwnd` must be a valid HWND with a window procedure that accepts
    /// the message `caller_msg` and produces no unsound behavior from doing so
    /// at any point in time.
    pub unsafe fn set_message(&self, caller_hwnd: HWND, caller_msg: u32) {
        let mut timer_mut = self.0.lock().unwrap();

        timer_mut.caller_hwnd = caller_hwnd;
        timer_mut.caller_msg = caller_msg;
        timer_mut.last = Instant::now();
        timer_mut.drift = Duration::new(0, 0);
    }

    /// Determine if the timer is running.
    pub fn is_stopped(&self) -> bool {
        let timer = self.0.lock().unwrap();
        timer.timer.is_invalid() || timer.caller_hwnd.is_invalid()
    }

    /// Manually tick the timer.
    pub fn tick(&self) {
        let (hwnd, msg) = {
            let timer_inner = self.0.lock().unwrap();

            (timer_inner.caller_hwnd, timer_inner.caller_msg)
        };

        if !hwnd.is_invalid() {
            unsafe { PostMessageW(hwnd, msg, WPARAM(0), LPARAM(0)) };
        }
    }

    /// The system timer callback.
    ///
    /// # Safety
    ///
    /// Must only be called by the system timer queue. Implicitly owns a
    /// reference to the Rust-side `Timer` that must not be dropped until the
    /// timer expires.
    unsafe extern "system" fn timer_callback(lp_param: *mut c_void, _timer_fired: BOOLEAN) {
        // As explained in `window_class` we have to take care not to consume
        // the lp_param; so we clone it and forget the original.
        let timer_orig = Arc::from_raw(lp_param as *const Mutex<TimerData>);
        let timer = Self(timer_orig.clone());
        forget(timer_orig);

        timer.tick();
    }

    /// Determine if the timer has elapsed.
    ///
    /// This should be called in the window loop thread immediately after
    /// receiving the message that the timer was told to send. Any operations
    /// that might cause unexpected jitter, such as locking a `Mutex` or
    /// cloning an `Arc`, should occur before calling `elapsed` to ensure the
    /// best timing.
    pub fn elapsed<P>(&self, p: P)
    where
        P: FnOnce(),
    {
        let mut timer_mut = self.0.lock().unwrap();
        let next = Instant::now();
        let new_drift = next - timer_mut.last;
        timer_mut.drift += new_drift;
        timer_mut.last = next;

        let target_period = timer_mut.target_period;
        if timer_mut.drift > target_period {
            timer_mut.drift -= target_period;

            p()
        }

        if timer_mut.drift > target_period {
            eprintln!(
                "WARNING: Excess drift of {:?} exceeds target period of {:?}",
                timer_mut.drift, timer_mut.target_period
            );
        }
    }

    /// Stop the timer.
    pub fn stop(&self) {
        let timer = self.0.lock().unwrap().timer;
        if timer.is_invalid() {
            return;
        }

        unsafe {
            DeleteTimerQueueTimer(HANDLE(0), timer, HANDLE(0));
        }

        let mut timer_mut = self.0.lock().unwrap();
        timer_mut.caller_hwnd = HWND(0);
        timer_mut.caller_msg = 0;
        timer_mut.last = Instant::now();
        timer_mut.drift = Duration::new(0, 0);
    }
}
