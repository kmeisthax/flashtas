//! ActiveX bindings for Rust

pub mod bindings;
mod rrf_com;

pub use bindings::stdole::IDispatch;
pub use rrf_com::get_class_object_by_dll;
