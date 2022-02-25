//! Exported from COM metadata to Rust via comexport-rs
//!
//! GUID: D27CDB6B-AE6D-11CF-96B8-444553540000
//! Version: 1.0
//! Locale ID: 0
//! SYSKIND: 3
//! Flags: 8

#![allow(clippy::too_many_arguments)]
#![allow(clippy::upper_case_acronyms)]

use crate::IDispatch;
use com::interfaces::IUnknown;
use com::sys::GUID;
use windows::core::HRESULT;
use windows::Win32::System::Com::{DISPPARAMS, EXCEPINFO};

type BSTR = *const u16;

/// Shockwave Flash
///
/// GUID: D27CDB6E-AE6D-11CF-96B8-444553540000
pub const SHOCKWAVE_FLASH_CLSID: GUID = GUID {
    data1: 0xD27CDB6E,
    data2: 0xAE6D,
    data3: 0x11CF,
    data4: [0x96, 0xB8, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0],
};

/// IFlashObjectInterface Interface
///
/// GUID: D27CDB71-AE6D-11CF-96B8-444553540000
pub const FLASH_OBJECT_INTERFACE_CLSID: GUID = GUID {
    data1: 0xD27CDB71,
    data2: 0xAE6D,
    data3: 0x11CF,
    data4: [0x96, 0xB8, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0],
};

/// FlashObject Class
///
/// GUID: E0920E11-6B65-4D5D-9C58-B1FC5C07DC43
pub const FLASH_OBJECT_CLSID: GUID = GUID {
    data1: 0xE0920E11,
    data2: 0x6B65,
    data3: 0x4D5D,
    data4: [0x9C, 0x58, 0xB1, 0xFC, 0x5C, 0x7, 0xDC, 0x43],
};

com::interfaces! {
    /// Shockwave Flash
    #[uuid("D27CDB6C-AE6D-11CF-96B8-444553540000")]
    pub unsafe interface IShockwaveFlash: IDispatch {
        // TODO: dispatch
        // fn QueryInterface(&self, param0: *mut GUID, param1: *mut *mut c_void);

        // TODO: dispatch
        // fn AddRef(&self) -> u32;

        // TODO: dispatch
        // fn Release(&self) -> u32;

        // TODO: dispatch
        // fn GetTypeInfoCount(&self, param0: *mut usize);

        // TODO: dispatch
        // fn GetTypeInfo(&self, param0: usize, param1: u32, param2: *mut *mut c_void);

        // TODO: dispatch
        // fn GetIDsOfNames(&self, param0: *mut GUID, param1: *mut *mut i8, param2: usize, param3: u32, param4: *mut i32);

        // TODO: dispatch
        // fn Invoke(&self, param0: i32, param1: *mut GUID, param2: u32, param3: u16, param4: *mut DISPPARAMS, param5: *mut IShockwaveFlash, param6: *mut EXCEPINFO, param7: *mut usize);

        //TODO: ReadyState (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: TotalFrames (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: Playing (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: Playing (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: Quality (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: Quality (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: ScaleMode (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: ScaleMode (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: AlignMode (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: AlignMode (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: BackgroundColor (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: BackgroundColor (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: Loop (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: Loop (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: Movie (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: Movie (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: FrameNum (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: FrameNum (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        // TODO: dispatch
        // fn SetZoomRect(&self, param0: i32, param1: i32, param2: i32, param3: i32);

        // TODO: dispatch
        // fn Zoom(&self, param0: isize);

        // TODO: dispatch
        // fn Pan(&self, param0: i32, param1: i32, param2: isize);

        // TODO: dispatch
        // fn Play(&self);

        // TODO: dispatch
        // fn Stop(&self);

        // TODO: dispatch
        // fn Back(&self);

        // TODO: dispatch
        // fn Forward(&self);

        // TODO: dispatch
        // fn Rewind(&self);

        // TODO: dispatch
        // fn StopPlay(&self);

        // TODO: dispatch
        // fn GotoFrame(&self, param0: i32);

        // TODO: dispatch
        // fn CurrentFrame(&self) -> i32;

        // TODO: dispatch
        // fn IsPlaying(&self) -> BOOL;

        // TODO: dispatch
        // fn PercentLoaded(&self) -> i32;

        // TODO: dispatch
        // fn FrameLoaded(&self, param0: i32) -> BOOL;

        // TODO: dispatch
        // fn FlashVersion(&self) -> i32;

        //TODO: WMode (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: WMode (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: SAlign (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: SAlign (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: Menu (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: Menu (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: Base (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: Base (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: Scale (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: Scale (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: DeviceFont (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: DeviceFont (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: EmbedMovie (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: EmbedMovie (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: BGColor (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: BGColor (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: Quality2 (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: Quality2 (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        // TODO: dispatch
        // fn LoadMovie(&self, param0: isize, param1: BSTR);

        // TODO: dispatch
        // fn TGotoFrame(&self, param0: BSTR, param1: i32);

        // TODO: dispatch
        // fn TGotoLabel(&self, param0: BSTR, param1: BSTR);

        // TODO: dispatch
        // fn TCurrentFrame(&self, param0: BSTR) -> i32;

        // TODO: dispatch
        // fn TCurrentLabel(&self, param0: BSTR) -> BSTR;

        // TODO: dispatch
        // fn TPlay(&self, param0: BSTR);

        // TODO: dispatch
        // fn TStopPlay(&self, param0: BSTR);

        // TODO: dispatch
        // fn SetVariable(&self, param0: BSTR, param1: BSTR);

        // TODO: dispatch
        // fn GetVariable(&self, param0: BSTR) -> BSTR;

        // TODO: dispatch
        // fn TSetProperty(&self, param0: BSTR, param1: isize, param2: BSTR);

        // TODO: dispatch
        // fn TGetProperty(&self, param0: BSTR, param1: isize) -> BSTR;

        // TODO: dispatch
        // fn TCallFrame(&self, param0: BSTR, param1: isize);

        // TODO: dispatch
        // fn TCallLabel(&self, param0: BSTR, param1: BSTR);

        // TODO: dispatch
        // fn TSetPropertyNum(&self, param0: BSTR, param1: isize, param2: f64);

        // TODO: dispatch
        // fn TGetPropertyNum(&self, param0: BSTR, param1: isize) -> f64;

        // TODO: dispatch
        // fn TGetPropertyAsNumber(&self, param0: BSTR, param1: isize) -> f64;

        //TODO: SWRemote (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: SWRemote (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: FlashVars (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: FlashVars (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: AllowScriptAccess (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: AllowScriptAccess (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: MovieData (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: MovieData (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: InlineData (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: InlineData (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: SeamlessTabbing (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: SeamlessTabbing (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        // TODO: dispatch
        // fn EnforceLocalSecurity(&self);

        //TODO: Profile (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: Profile (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: ProfileAddress (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: ProfileAddress (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: ProfilePort (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: ProfilePort (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        // TODO: dispatch
        // fn CallFunction(&self, param0: BSTR) -> BSTR;

        // TODO: dispatch
        // fn SetReturnValue(&self, param0: BSTR);

        // TODO: dispatch
        // fn DisableLocalSecurity(&self);

        //TODO: AllowNetworking (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: AllowNetworking (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: AllowFullScreen (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: AllowFullScreen (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: AllowFullScreenInteractive (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: AllowFullScreenInteractive (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: IsDependent (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: IsDependent (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: BrowserZoom (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: BrowserZoom (funckind FUNCKIND(4), invkind INVOKEKIND(4))
        //TODO: IsTainted (funckind FUNCKIND(4), invkind INVOKEKIND(2))
        //TODO: IsTainted (funckind FUNCKIND(4), invkind INVOKEKIND(4))
    }

    /// Event interface for Shockwave Flash
    #[uuid("D27CDB6D-AE6D-11CF-96B8-444553540000")]
    pub unsafe interface _IShockwaveFlashEvents: IDispatch {
        // TODO: dispatch
        // fn OnReadyStateChange(&self, param0: i32);

        // TODO: dispatch
        // fn OnProgress(&self, param0: i32);

        // TODO: dispatch
        // fn FSCommand(&self, param0: BSTR, param1: BSTR);

        // TODO: dispatch
        // fn FlashCall(&self, param0: BSTR);

    }

    /// IFlashFactory Interface
    #[uuid("D27CDB70-AE6D-11CF-96B8-444553540000")]
    pub unsafe interface IFlashFactory: IUnknown {
    }

    /// IFlashObjectInterface Interface
    #[uuid("D27CDB72-AE6D-11CF-96B8-444553540000")]
    pub unsafe interface IFlashObjectInterface: IDispatchEx {
    }

    #[uuid("A6EF9860-C720-11D0-9337-00A0C90DCAA9")]
    pub unsafe interface IDispatchEx: IDispatch {
        fn GetDispID(&self, param0: BSTR, param1: u32, param2: *mut i32) -> HRESULT;
        fn RemoteInvokeEx(&self, param0: i32, param1: u32, param2: u32, param3: *mut DISPPARAMS, param4: *mut IShockwaveFlash, param5: *mut EXCEPINFO, param6: *mut IServiceProvider, param7: usize, param8: *mut usize, param9: *mut IShockwaveFlash) -> HRESULT;
        fn DeleteMemberByName(&self, param0: BSTR, param1: u32) -> HRESULT;
        fn DeleteMemberByDispID(&self, param0: i32) -> HRESULT;
        fn GetMemberProperties(&self, param0: i32, param1: u32, param2: *mut u32) -> HRESULT;
        fn GetMemberName(&self, param0: i32, param1: *mut BSTR) -> HRESULT;
        fn GetNextDispID(&self, param0: u32, param1: i32, param2: *mut i32) -> HRESULT;
        fn GetNameSpaceParent(&self, param0: *mut IUnknown) -> HRESULT;
    }

    #[uuid("6D5140C1-7436-11CE-8034-00AA006009FA")]
    pub unsafe interface IServiceProvider: IUnknown {
        fn RemoteQueryService(&self, param0: *mut GUID, param1: *mut GUID, param2: *mut IUnknown) -> HRESULT;
    }

    /// IFlashObject Interface
    #[uuid("86230738-D762-4C50-A2DE-A753E5B1686F")]
    pub unsafe interface IFlashObject: IDispatch {
        // TODO: dispatch
        // fn QueryInterface(&self, param0: *mut GUID, param1: *mut *mut c_void);

        // TODO: dispatch
        // fn AddRef(&self) -> u32;

        // TODO: dispatch
        // fn Release(&self) -> u32;

        // TODO: dispatch
        // fn GetTypeInfoCount(&self, param0: *mut usize);

        // TODO: dispatch
        // fn GetTypeInfo(&self, param0: usize, param1: u32, param2: *mut *mut c_void);

        // TODO: dispatch
        // fn GetIDsOfNames(&self, param0: *mut GUID, param1: *mut *mut i8, param2: usize, param3: u32, param4: *mut i32);

        // TODO: dispatch
        // fn Invoke(&self, param0: i32, param1: *mut GUID, param2: u32, param3: u16, param4: *mut DISPPARAMS, param5: *mut IShockwaveFlash, param6: *mut EXCEPINFO, param7: *mut usize);

        // TODO: dispatch
        // fn GetDispID(&self, param0: BSTR, param1: u32, param2: *mut i32);

        // TODO: dispatch
        // fn RemoteInvokeEx(&self, param0: i32, param1: u32, param2: u32, param3: *mut DISPPARAMS, param4: *mut IShockwaveFlash, param5: *mut EXCEPINFO, param6: *mut IServiceProvider, param7: usize, param8: *mut usize, param9: *mut IShockwaveFlash);

        // TODO: dispatch
        // fn DeleteMemberByName(&self, param0: BSTR, param1: u32);

        // TODO: dispatch
        // fn DeleteMemberByDispID(&self, param0: i32);

        // TODO: dispatch
        // fn GetMemberProperties(&self, param0: i32, param1: u32, param2: *mut u32);

        // TODO: dispatch
        // fn GetMemberName(&self, param0: i32, param1: *mut BSTR);

        // TODO: dispatch
        // fn GetNextDispID(&self, param0: u32, param1: i32, param2: *mut i32);

        // TODO: dispatch
        // fn GetNameSpaceParent(&self, param0: *mut IUnknown);

    }

}
