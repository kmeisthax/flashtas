//! ActiveX bindings for Rust

mod dispatch;
pub mod flash;
mod rrf_com;

pub use dispatch::IDispatch;
pub use rrf_com::get_class_object_by_dll;
