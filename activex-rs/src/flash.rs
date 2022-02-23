//! Exported from COM metadata to Rust via comexport-rs
//!
//! GUID: D27CDB6B-AE6D-11CF-96B8-444553540000
//! Version: 1.0
//! Locale ID: 0
//! SYSKIND: 3
//! Flags: 8

use com::interfaces::IUnknown;
use com::sys::GUID;

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
    pub unsafe interface IShockwaveFlash: IUnknown {
        //TODO: QueryInterface
        //TODO: AddRef
        //TODO: Release
        //TODO: GetTypeInfoCount
        //TODO: GetTypeInfo
        //TODO: GetIDsOfNames
        //TODO: Invoke
        //TODO: ReadyState
        //TODO: TotalFrames
        //TODO: Playing
        //TODO: Playing
        //TODO: Quality
        //TODO: Quality
        //TODO: ScaleMode
        //TODO: ScaleMode
        //TODO: AlignMode
        //TODO: AlignMode
        //TODO: BackgroundColor
        //TODO: BackgroundColor
        //TODO: Loop
        //TODO: Loop
        //TODO: Movie
        //TODO: Movie
        //TODO: FrameNum
        //TODO: FrameNum
        //TODO: SetZoomRect
        //TODO: Zoom
        //TODO: Pan
        //TODO: Play
        //TODO: Stop
        //TODO: Back
        //TODO: Forward
        //TODO: Rewind
        //TODO: StopPlay
        //TODO: GotoFrame
        //TODO: CurrentFrame
        //TODO: IsPlaying
        //TODO: PercentLoaded
        //TODO: FrameLoaded
        //TODO: FlashVersion
        //TODO: WMode
        //TODO: WMode
        //TODO: SAlign
        //TODO: SAlign
        //TODO: Menu
        //TODO: Menu
        //TODO: Base
        //TODO: Base
        //TODO: Scale
        //TODO: Scale
        //TODO: DeviceFont
        //TODO: DeviceFont
        //TODO: EmbedMovie
        //TODO: EmbedMovie
        //TODO: BGColor
        //TODO: BGColor
        //TODO: Quality2
        //TODO: Quality2
        //TODO: LoadMovie
        //TODO: TGotoFrame
        //TODO: TGotoLabel
        //TODO: TCurrentFrame
        //TODO: TCurrentLabel
        //TODO: TPlay
        //TODO: TStopPlay
        //TODO: SetVariable
        //TODO: GetVariable
        //TODO: TSetProperty
        //TODO: TGetProperty
        //TODO: TCallFrame
        //TODO: TCallLabel
        //TODO: TSetPropertyNum
        //TODO: TGetPropertyNum
        //TODO: TGetPropertyAsNumber
        //TODO: SWRemote
        //TODO: SWRemote
        //TODO: FlashVars
        //TODO: FlashVars
        //TODO: AllowScriptAccess
        //TODO: AllowScriptAccess
        //TODO: MovieData
        //TODO: MovieData
        //TODO: InlineData
        //TODO: InlineData
        //TODO: SeamlessTabbing
        //TODO: SeamlessTabbing
        //TODO: EnforceLocalSecurity
        //TODO: Profile
        //TODO: Profile
        //TODO: ProfileAddress
        //TODO: ProfileAddress
        //TODO: ProfilePort
        //TODO: ProfilePort
        //TODO: CallFunction
        //TODO: SetReturnValue
        //TODO: DisableLocalSecurity
        //TODO: AllowNetworking
        //TODO: AllowNetworking
        //TODO: AllowFullScreen
        //TODO: AllowFullScreen
        //TODO: AllowFullScreenInteractive
        //TODO: AllowFullScreenInteractive
        //TODO: IsDependent
        //TODO: IsDependent
        //TODO: BrowserZoom
        //TODO: BrowserZoom
        //TODO: IsTainted
        //TODO: IsTainted
    }

    /// Event interface for Shockwave Flash
    #[uuid("D27CDB6D-AE6D-11CF-96B8-444553540000")]
    pub unsafe interface _IShockwaveFlashEvents: IUnknown {
        //TODO: OnReadyStateChange
        //TODO: OnProgress
        //TODO: FSCommand
        //TODO: FlashCall
    }

    /// IFlashFactory Interface
    #[uuid("D27CDB70-AE6D-11CF-96B8-444553540000")]
    pub unsafe interface IFlashFactory: IUnknown {
    }

    /// IFlashObjectInterface Interface
    #[uuid("D27CDB72-AE6D-11CF-96B8-444553540000")]
    pub unsafe interface IFlashObjectInterface: IUnknown {
    }

    #[uuid("A6EF9860-C720-11D0-9337-00A0C90DCAA9")]
    pub unsafe interface IDispatchEx: IUnknown {
        //TODO: GetDispID
        //TODO: RemoteInvokeEx
        //TODO: DeleteMemberByName
        //TODO: DeleteMemberByDispID
        //TODO: GetMemberProperties
        //TODO: GetMemberName
        //TODO: GetNextDispID
        //TODO: GetNameSpaceParent
    }

    #[uuid("6D5140C1-7436-11CE-8034-00AA006009FA")]
    pub unsafe interface IServiceProvider: IUnknown {
        //TODO: RemoteQueryService
    }

    /// IFlashObject Interface
    #[uuid("86230738-D762-4C50-A2DE-A753E5B1686F")]
    pub unsafe interface IFlashObject: IUnknown {
        //TODO: QueryInterface
        //TODO: AddRef
        //TODO: Release
        //TODO: GetTypeInfoCount
        //TODO: GetTypeInfo
        //TODO: GetIDsOfNames
        //TODO: Invoke
        //TODO: GetDispID
        //TODO: RemoteInvokeEx
        //TODO: DeleteMemberByName
        //TODO: DeleteMemberByDispID
        //TODO: GetMemberProperties
        //TODO: GetMemberName
        //TODO: GetNextDispID
        //TODO: GetNameSpaceParent
    }

}
