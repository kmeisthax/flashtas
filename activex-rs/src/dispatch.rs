//! IDispatch bindings

use com::interfaces::IUnknown;

com::interfaces! {
    #[uuid("00020400-0000-0000-C000-000000000046")]
    pub unsafe interface IDispatch: IUnknown {

    }
}
