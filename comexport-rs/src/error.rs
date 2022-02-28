use std::fmt::Error as FmtError;
use windows::core::Error as WinError;

#[derive(Debug)]
pub enum Error {
    Windows(WinError),
    RustFmt(FmtError),
    NoFuncDescForComFn,
    NoTypeAttrForComType,
    NoVarDescForComVar,
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
