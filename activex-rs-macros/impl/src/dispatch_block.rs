//! Dispatch block tree type

use proc_macro2::Ident;
use std::collections::HashMap;

/// A pair of identifier and invokekind for a dispatch method.
pub struct DispIdentifier(Ident);

/// A dispatch block.
/// 
/// Intended syntax is:
/// 
/// ```activex_rs::dispatch!(COClass) {
///     fn Method1(&self, param: i32) -> i32 {
///         param * 2
///     }
/// 
///     fn get Property(&self) -> i32 {
///         256
///     }
///     
///     #[disp(5)]
///     fn put Property(&self, val: i32) {
///         //TODO: Do something with the property.
///     }
/// }```
pub struct DispatchBlock {
    coclass: Ident,
    impls: HashMap<DispIdentifier, ()>
}