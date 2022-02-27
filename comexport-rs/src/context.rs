//! Code generation context

use crate::type_bridge::BridgedTypeLibrary;

/// Owned version of `Context`
#[derive(Default)]
pub struct CodeGen {
    types: BridgedTypeLibrary,
    structs: String,
    interfaces: String,
    impls: String,
}

impl CodeGen {
    pub fn new() -> Self {
        Self {
            types: BridgedTypeLibrary::with_predefined_bridges(),
            structs: String::new(),
            interfaces: String::new(),
            impls: String::new(),
        }
    }

    pub fn borrow(&mut self) -> Context<'_> {
        Context {
            types: &mut self.types,
            structs: &mut self.structs,
            interfaces: &mut self.interfaces,
            impls: &mut self.impls,
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

    /// Codegen stream for struct and class definitions before the COM
    /// interfaces block.
    pub structs: &'a mut String,

    /// Codegen stream for the COM interfaces block.
    ///
    /// Code written here will be wrapped inside of `com::interfaces!`.
    pub interfaces: &'a mut String,

    /// Codegen stream for `impl` blocks after the COM interfaces block.
    pub impls: &'a mut String,
}
