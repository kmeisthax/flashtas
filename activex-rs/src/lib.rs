//! ActiveX bindings for Rust

pub mod bindings;
mod ole_automation;
mod rrf_com;

pub use bindings::stdole::{IDispatch, IMoniker, IOleContainer};
pub use rrf_com::get_class_object_by_dll;
