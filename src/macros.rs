//use result::{ClrError, SourceLocation};
/*match hr {
    0 => Ok(()), 
    _ => Err(ClrError::InnerCall{hr: hr, source: SourceLocation::ICollection{line: line!()}})
}*/
#[macro_export]
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
#[macro_export]
macro_rules! PROPERTY {
    ($fn_name:ident $err_type:ident {get {$prop_name:ident (VARIANT_BOOL)}}) => {
        fn $prop_name(&self) -> Result<bool> 
        {
            let p = self.ptr_mut();
            let mut vb: *mut VARIANT_BOOL = ptr::null_mut();
            let hr = unsafe {
                (*p).$fn_name(&mut vb)
            };
            SUCCEEDED!(hr, unsafe{*vb} < 0, $err_type)
        }
    };

    ($fn_name:ident $err_type:ident {get {$prop_name:ident ($ptr_type:ident)}}) => {
        fn $prop_name<T>(&self) -> Result<T> 
            where T: PtrContainer<$ptr_type>
        {
            let p = self.ptr_mut();
            let mut t: *mut $ptr_type = ptr::null_mut();
            let hr = unsafe {
                (*p).$fn_name(&mut t)
            };
            SUCCEEDED!(hr, T::from(t), $err_type)
        }
    };
}

#[macro_export]
macro_rules! EXTRACT_VECTOR_FROM_SAFEARRAY {
    ($enum_name:ident, $psa_name:ident, $origin_type:ty, $transmuted_type:ty, $ctr_type:ident) => {
        {
            let rsa: RSafeArray<$ctr_type> = RSafeArray::from($psa_name);
            if let RSafeArray::$enum_name(array, _) = rsa {
                array.into_iter().map(|item| {
                    let trans_item = unsafe { mem::transmute::<$origin_type, $transmuted_type>(item)};
                    $ctr_type::from(trans_item)
                }).collect()
            }
            else {
                Vec::new()
            }
        }
        
    };
}

#[macro_export]
macro_rules! SIMPLE_EXTRACT {
    ($enum_name:ident, $psa_name:ident, $ptr_type:ty) => {
        {
            let rsa: RSafeArray<$ptr_type> = RSafeArray::from($psa_name);
            if let RSafeArray::$enum_name(inner) = rsa {
                inner
            } else {
                Vec::new()
            }
        }
    };
}