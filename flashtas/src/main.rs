use activex_rs::bindings::flash::{IShockwaveFlash, SHOCKWAVE_FLASH_CLSID};
use com::runtime::{create_instance, init_runtime};

fn main() {
    init_runtime().expect("A working COM runtime");

    let fp = create_instance::<IShockwaveFlash>(&SHOCKWAVE_FLASH_CLSID).expect("Flash");

    unsafe { fp.Zoom(1).expect("No whammies no whammies no whammies") };

    println!(
        "Flash version: {}",
        unsafe { fp.FlashVersion() }.expect("Flash version")
    );
}
