//! Code generation context

use crate::type_bridge::BridgedTypeLibrary;

#[derive(Default)]
pub struct CodeGen {
    types: BridgedTypeLibrary,
}

impl CodeGen {
    pub fn new() -> Self {
        Self {
            types: BridgedTypeLibrary::with_predefined_bridges(),
        }
    }

    pub fn borrow(&mut self) -> Context<'_> {
        Context {
            types: &mut self.types,
        }
    }
}

/// The current codegen context.
///
/// This should be passed around all functions that do code generation, in lieu
/// of explicit function parameters and string returns.
pub struct Context<'a> {
    /// Universe of all known types.
    pub types: &'a mut BridgedTypeLibrary,
}
