use activex_rs::bindings::flash::{IShockwaveFlash, SHOCKWAVE_FLASH_CLSID};
use activex_rs::get_class_object_by_dll;
use com::interfaces::IClassFactory;
use com::runtime::init_runtime;
use std::env::args;
use std::ffi::OsStr;

fn main() {
    init_runtime().expect("A working COM runtime");

    let path = args().nth(1).expect("Path to .exe, .dll, or .ocx");
    let fp_ocx_name = OsStr::new(&path);
    let fp_class = get_class_object_by_dll::<IClassFactory>(fp_ocx_name, &SHOCKWAVE_FLASH_CLSID)
        .expect("ShockwaveFlash class");
    let _fp = fp_class
        .create_instance::<IShockwaveFlash>()
        .expect("A working FlashPlayer");

    //Do something else, presumably.
}
