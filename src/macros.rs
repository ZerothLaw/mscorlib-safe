//use result::{ClrError, SourceLocation};
/*match hr {
    0 => Ok(()), 
    _ => Err(ClrError::InnerCall{hr: hr, source: SourceLocation::ICollection{line: line!()}})
}*/
macro_rules! SUCCEEDED {
    ($hr:ident, $ok:tt, $source:ident) => {
        match $hr {
            0 => Ok($ok), 
            _ => Err(ClrError::InnerCall{hr: $hr, source: SourceLocation::$source(line!())})
        }
    };
    ($hr:ident, $ok:expr, $source:ident) => {
        match $hr {
            0 => Ok($ok), 
            _ => Err(ClrError::InnerCall{hr: $hr, source: SourceLocation::$source(line!())})
        }
    };
}
//PROPERTY!{DeclaringType {
//  get{
//      declaring_type(_Type)
//  }} 
//}
/*
    fn declaring_type<T>(&self) -> Result<T> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let t: *mut *mut _Type = ptr::null_mut();
        let hr = unsafe {
            (*p).get_DeclaringType(t)
        };
        SUCCEEDED!(hr, T::from(unsafe{*t}), _Type)
    }

    fn $prop_name<T>(&self) -> Result<T> 
        where T: PtrContainer<$ptr_type>
    {
        let p = self.ptr_mut();
        let t: *mut *mut $ptr_type = ptr::null_mut();
        let hr = unsafe {
            (*p).$prop_type_$fn_name(t)
        };
        SUCCEEDED!(hr, T::from(unsafe{*t}), $ptr_type)
    }
    PROPERTY!{get_IsNotPublic _Type{ get { not_public(VARIANT_BOOL) }}}
*/
macro_rules! PROPERTY {
    ($fn_name:ident $err_type:ident {get {$prop_name:ident (VARIANT_BOOL)}}) => {
        fn $prop_name(&self) -> Result<bool> 
        {
            let p = self.ptr_mut();
            let t: *mut *mut VARIANT_BOOL = ptr::null_mut();
            let hr = unsafe {
                (*p).$fn_name(t)
            };
            SUCCEEDED!(hr, unsafe{**t} > 0, $err_type)
        }
    };

    ($fn_name:ident $err_type:ident {get {$prop_name:ident ($ptr_type:ident)}}) => {
        fn $prop_name<T>(&self) -> Result<T> 
            where T: PtrContainer<$ptr_type>
        {
            let p = self.ptr_mut();
            let t: *mut *mut $ptr_type = ptr::null_mut();
            let hr = unsafe {
                (*p).$fn_name(t)
            };
            SUCCEEDED!(hr, T::from(unsafe{*t}), $err_type)
        }
    };
}