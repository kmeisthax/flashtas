use activex_rs::bindings::ole32::IOleClientSite;
use activex_rs::bindings::stdole::{IMoniker, IOleContainer};
use com::production::ClassAllocation;
use lazy_static::lazy_static;
use windows::core::HRESULT;

com::class! {
    /// OLE client site for FlashTAS client controls
    pub class TASClientSite: IOleClientSite {

    }

    impl IOleClientSite for TASClientSite {
        fn SaveObject(&self) -> HRESULT {
            unimplemented!();
        }

        fn GetMoniker(&self, param0: u32, param1: u32, param2: *mut IMoniker) -> HRESULT {
            unimplemented!();
        }

        fn GetContainer(&self, param0: *mut IOleContainer) -> HRESULT {
            unimplemented!();
        }

        fn ShowObject(&self) -> HRESULT {
            unimplemented!();
        }

        fn OnShowWindow(&self, param0: ::com::sys::BOOL) -> HRESULT {
            unimplemented!();
        }

        fn RequestNewObjectLayout(&self) -> HRESULT {
            unimplemented!();
        }
    }
}

lazy_static! {
    pub static ref TASClientSite__CF: ClassAllocation<TASClientSiteClassFactory> =
        TASClientSiteClassFactory::allocate();
}
