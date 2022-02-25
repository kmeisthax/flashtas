//! ActiveX bindings for Rust

pub mod flash;
mod rrf_com;
mod stdole;

pub use rrf_com::get_class_object_by_dll;
pub use stdole::IDispatch;
