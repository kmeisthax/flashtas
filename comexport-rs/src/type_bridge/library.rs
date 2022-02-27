use com::interfaces::{IClassFactory, IUnknown};
use com::sys::GUID as ComGUID;
use com::Interface;
use std::borrow::Cow;
use std::collections::HashMap;
use std::mem::transmute;
use windows::core::GUID;

/// Convert from COM to Windows `GUID` structures.
///
/// Both structures are `repr(C)` and have identical field layout; thus it
/// should be safe to transmute between them.
///
/// TODO: Remove if `com` removes it's `GUID` type.
pub trait ToWindowsGuid {
    fn to_win_guid(self) -> GUID;
}

impl ToWindowsGuid for GUID {
    fn to_win_guid(self) -> GUID {
        self
    }
}

impl ToWindowsGuid for ComGUID {
    fn to_win_guid(self) -> GUID {
        unsafe { transmute(self) }
    }
}

/// Information about an encountered type.
pub struct BridgedType {
    /// The type's ID.
    ///
    /// If `None`, the type does not have a valid GUID (all zeroes). COM types
    /// only have a valid GUID if they are coclasses or interfaces.
    id: Option<GUID>,

    /// The Rust-side name of the type.
    rust_name: Cow<'static, str>,

    /// The Ruse-side module that the type is defined in.
    ///
    /// `None` indicates types in the current module.
    rust_module: Option<&'static str>,
}

impl BridgedType {
    /// Pre-defined a type as already bridged.
    ///
    /// This is useful for declaring certain types as already having been
    /// bridged by other modules. For example, `IUnknown` is bridged by
    /// `com-rs`; type libraries that redefine it do not need to be bridged
    /// again.
    fn from_predefined(
        id: Option<GUID>,
        rust_name: impl Into<Cow<'static, str>>,
        rust_module: &'static str,
    ) -> Self {
        Self {
            id,
            rust_name: rust_name.into(),
            rust_module: Some(rust_module),
        }
    }

    /// Create a type by bridging it in the current module
    fn from_define(guid: impl ToWindowsGuid, rust_name: impl Into<Cow<'static, str>>) -> Self {
        let guid = guid.to_win_guid();
        let id = if guid == GUID::zeroed() {
            None
        } else {
            Some(guid)
        };

        Self {
            id,
            rust_name: rust_name.into(),
            rust_module: None,
        }
    }

    pub fn rust_name(&self) -> &str {
        &self.rust_name
    }
}

#[derive(Default)]
pub struct BridgedTypeLibrary {
    defs: Vec<BridgedType>,

    /// Index of all GUID-bearing bridged types.
    guid_idx: HashMap<GUID, usize>,

    /// Index of all named bridged types.
    name_idx: HashMap<Cow<'static, str>, usize>,
}

impl BridgedTypeLibrary {
    /// Construct a new type library with bridging definitions for known
    /// classes.
    pub fn with_predefined_bridges() -> Self {
        let mut lib = Self::default();

        lib.predef_i(IUnknown::IID, "IUnknown", "com::interfaces");
        lib.predef_i(IClassFactory::IID, "IClassFactory", "com::interfaces");
        lib.predef_i(
            GUID::from_u128(0x00020400_0000_0000_C000_000000000046),
            "IDispatch",
            "crate",
        );
        lib.predef_s("GUID", "com::sys");
        lib.predef_s("HRESULT", "windows::core");
        lib.predef_s("BOOL", "windows::win32::Foundation::BOOL");
        lib.predef_s("DISPPARAMS", "windows::win32::System::Com::DISPPARAMS");
        lib.predef_s("EXCEPINFO", "windows::win32::System::Com::EXCEPINFO");
        lib.predef_s("SAFEARRAY", "windows::win32::System::Com::SAFEARRAY");
        lib.predef_s("VARIANT", "windows::win32::System::Com::VARIANT");
        lib.predef_s(
            "DISPATCH_METHOD",
            "windows::win32::System::Ole::DISPATCH_METHOD",
        );
        lib.predef_s("VARENUM", "windows::win32::System::Ole::VARENUM");

        lib
    }

    fn define_bridged_type(&mut self, bridged_type: BridgedType) -> usize {
        let index = self.defs.len();
        let id = bridged_type.id;
        let name = bridged_type.rust_name.clone();

        self.defs.push(bridged_type);

        if let Some(id) = id {
            self.guid_idx.insert(id, index);
        }

        self.name_idx.insert(name, index);

        index
    }

    /// Predefine a COM interface of a given ID and name.
    fn predef_i(
        &mut self,
        id: impl ToWindowsGuid,
        rust_name: impl Clone + Into<Cow<'static, str>>,
        rust_module: &'static str,
    ) {
        let win_id = id.to_win_guid();
        self.define_bridged_type(BridgedType::from_predefined(
            Some(win_id),
            rust_name,
            rust_module,
        ));
    }

    /// Predefine a COM structure of a given name.
    fn predef_s(
        &mut self,
        rust_name: impl Clone + Into<Cow<'static, str>>,
        rust_module: &'static str,
    ) {
        self.define_bridged_type(BridgedType::from_predefined(None, rust_name, rust_module));
    }

    /// Retrieve a bridged type by it's GUID.
    pub fn type_by_guid(&self, guid: impl ToWindowsGuid) -> Option<&BridgedType> {
        let guid = guid.to_win_guid();
        if guid == GUID::zeroed() {
            return None;
        }

        self.guid_idx.get(&guid).and_then(|i| self.defs.get(*i))
    }

    /// Retrieve a bridged type by it's Rust-side name.
    pub fn type_by_name(&self, name: &str) -> Option<&BridgedType> {
        self.name_idx.get(name).and_then(|i| self.defs.get(*i))
    }

    /// Bridge a struct or class by defining it in the current module.
    pub fn define_generated_bridge(&mut self, guid: impl ToWindowsGuid, name: String) {
        self.define_bridged_type(BridgedType::from_define(guid, name));
    }
}
