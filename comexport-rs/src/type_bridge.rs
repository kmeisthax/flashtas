//! Rust/COM type bridging helpers
//!
//! Functions in this package are responsible for taking a COM type and picking
//! a suitable Rust equivalent that matches.

mod library;
mod typedesc;

pub use library::BridgedTypeLibrary;
pub use typedesc::{bridge_elem_to_rust_type, bridge_usertype_to_rust_type, bridged_hreftype};
