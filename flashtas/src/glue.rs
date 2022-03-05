//! Annoying glue logic to deal with various crates' types

/// Write a COM pointer to an output type.
///
/// For some reason, directly writing owned COM types through a pointer to that
/// type causes an immediate access violation. This somehow gets around that by
/// transmuting away the COM types and just working with pointers to `c_void`.
///
/// # Safety
///
/// `val` must be a valid COM interface pointer. `param` must be a valid
/// pointer or reference to somewhere that can hold a COM interface pointer.
///
/// Reference counts are not adjusted by this function. Thus, this function
/// transfers ownership of the reference from `val` to `*param`.
///
/// `T` and `U` must be the same underlying COM interface. Competing bridge
/// types for those interfaces may be mixed and matched otherwise. For example,
/// it is valid to use `write_com` to bridge a `windows::core::IUnknown` to a
/// `com::interfaces::IUnknown`.
#[macro_export]
macro_rules! unsafe_write_com {
    ($param:ident, $val:ident) => {
        #[allow(unused_unsafe)]
        unsafe {
            let inner_param = ::std::mem::transmute::<_, *mut *mut ::std::ffi::c_void>($param);
            let inner_val = ::std::mem::transmute::<_, *mut ::std::ffi::c_void>($val);

            *inner_param = inner_val;
        }
    };
}
