use std::fmt::Error as FmtError;
use windows::core::Error as WinError;
use windows::Win32::System::Ole::VARENUM;

#[derive(Debug)]
pub enum Error {
    Windows(WinError),
    RustFmt(FmtError),
    NoFuncDescForComFn,
    NoTypeAttrForComType,
    NoVarDescForComVar,
    UnknownVariantType(VARENUM),
}

impl From<WinError> for Error {
    fn from(error: WinError) -> Error {
        Error::Windows(error)
    }
}

impl From<FmtError> for Error {
    fn from(error: FmtError) -> Error {
        Error::RustFmt(error)
    }
}

impl From<VARENUM> for Error {
    fn from(error: VARENUM) -> Error {
        Error::UnknownVariantType(error)
    }
}
