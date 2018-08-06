use std::mem;
use std::ptr;

use winapi::ctypes::c_long; 
use winapi::shared::wtypes::VARIANT_BOOL;
use winapi::um::oaidl::{IDispatch, VARIANT};

use mscorlib_sys::system::{IComparable, _Array};
use mscorlib_sys::system::collections::{DictionaryEntry, ICollection, IComparer, IDictionary, IDictionaryEnumerator, 
IEnumerable, IEnumerator, IEqualityComparer, IHashCodeProvider, IList};

use new_variant::Variant;
use wrappers::PtrContainer;
use result::{ClrError, SourceLocation, Result};

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
pub trait Collection where Self: PtrContainer<ICollection> {
    fn copy_to<R>(&self, index: i32, rhs: &R) -> Result<()>
        where R: PtrContainer<_Array>
    {
        let lhs_ptr: *mut ICollection = self.ptr_mut();
        let rhs_vt = rhs.ptr_mut();
        let hr = unsafe {
            (*lhs_ptr).CopyTo(rhs_vt, index)
        };
        SUCCEEDED!(hr, (), ICollection)
    }

    fn count(&self) -> Result<i32>{
        let mut p_c: c_long = 0;
        let p = self.ptr_mut();
        let hr = unsafe{
            (*p).get_Count(&mut p_c)
        };
        SUCCEEDED!(hr, p_c, ICollection)
    }

    fn synchronized(&self) -> Result<bool>{
        let mut vb: VARIANT_BOOL = 0;
        let p = self.ptr_mut();
        let hr = unsafe {
            (*p).get_IsSynchronized(&mut vb)
        };
        let b = vb < 0;
        SUCCEEDED!(hr, b, ICollection)
    }
}

pub trait Comparable where Self: PtrContainer<IComparable> {
    fn compare<R, T>(&self, rhs: R) -> Result<i32>
        where R: Comparable + PtrContainer<IComparable>
    {
        let lhs_ptr: *mut IComparable = self.ptr_mut();
        let p_dispatch = unsafe{mem::transmute::<*mut IComparable, *mut IDispatch>(rhs.ptr_mut())};
        let rhs_vt : VARIANT = VARIANT::from(Variant::from(p_dispatch));
        let mut ret: c_long = 0;
        let hr = unsafe {
            (*lhs_ptr).CompareTo(rhs_vt, &mut ret)
        };

        SUCCEEDED!(hr, ret, IComparable)
    }
}

pub trait Comparer where Self: PtrContainer<IComparer> {
    fn compare<L, R, TLeft, TRight>(&self, lhs: L, rhs: R) -> Result<i32>
        where L: PtrContainer<TLeft>, 
              R: PtrContainer<TRight>
    {
        let p = self.ptr_mut();
        let lhs_vt : VARIANT = VARIANT::from(lhs.into_variant());
        let rhs_vt : VARIANT = VARIANT::from(rhs.into_variant());
        let mut ret: c_long = 0;
        let hr = unsafe {
            (*p).Compare(lhs_vt, rhs_vt, &mut ret)
        };
        SUCCEEDED!(hr, ret, IComparer)
    }
}

pub trait Dictionary where Self: PtrContainer<IDictionary> {
    fn item<K, TDispatch, V, TDispatch2>(&self, key: K) -> Result<Variant>
        where K: PtrContainer<TDispatch>, 
              V: PtrContainer<TDispatch2>
    {
        let p = self.ptr_mut();
        let vt : VARIANT = VARIANT::from(key.into_variant());
        let mut ret: VARIANT = unsafe {mem::zeroed()};
        let hr = unsafe {
            (*p).get_Item(vt, &mut ret)
        };
        SUCCEEDED!(hr, Variant::from(ret), IDictionary)
    }

    fn item_mut<K, V, TDispatch, TDispatch2>(&mut self, key: K, value: V) -> Result<()>
        where K: PtrContainer<TDispatch>, 
              V: PtrContainer<TDispatch2>
    {
        let p = self.ptr_mut();
        let kvt : VARIANT = VARIANT::from(key.into_variant());
        let vvt : VARIANT = VARIANT::from(value.into_variant());
        let hr = unsafe {
            (*p).putref_Item(kvt, vvt)
        };
        SUCCEEDED!(hr, (), IDictionary)
    }

    fn keys<C: Collection>(&self) -> Result<C>
    {
        let p = self.ptr_mut();
        let mut ic: *mut ICollection = ptr::null_mut();
        let hr = unsafe {
            (*p).get_Keys(&mut ic)
        };
        SUCCEEDED!(hr, C::from(ic), IDictionary)
    }

    fn values<C: Collection>(&self) -> Result<C>
    {
        let p = self.ptr_mut();
        let mut pic: *mut ICollection = ptr::null_mut();
        let hr = unsafe {
            (*p).get_Values(&mut pic) 
        };
        SUCCEEDED!(hr, C::from(pic), IDictionary)
    }

    fn contains<TOut, V>(&self, obj: V) -> Result<bool>
        where V: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt : VARIANT = VARIANT::from(obj.into_variant());
        let mut pb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Contains(vt, &mut pb)
        };
        SUCCEEDED!(hr, pb < 0, IDictionary)
    }

    fn add<K, V, TKey, TValue>(&self, key: K, value: V) -> Result<()>
        where K: PtrContainer<TKey>,
              V: PtrContainer<TValue>
    {
        let k : VARIANT = VARIANT::from(key.into_variant());
        let v : VARIANT = VARIANT::from(value.into_variant());
        let p = self.ptr_mut();
        let hr = unsafe {
            (*p).Add(k, v)
        };
        SUCCEEDED!(hr, (), IDictionary)
    }
    
    fn clear(&self) -> Result<()>{
        let p = self.ptr_mut();
        let hr = unsafe {
            (*p).Clear()
        };
        SUCCEEDED!(hr, (), IDictionary)
    }

    fn read_only(&self) -> Result<bool>{
        let p = self.ptr_mut();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).get_IsReadOnly(&mut vb)
        };
        SUCCEEDED!(hr, vb < 0, IDictionary)
    }

    fn fixed_size(&self) -> Result<bool>{
        let p = self.ptr_mut();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).get_IsFixedSize(&mut vb)
        };
        SUCCEEDED!(hr, vb < 0, IDictionary)
    }

    fn enumerator<DE>(&self) -> Result<DE>
        where DE: From<*mut IDictionaryEnumerator>
    {
        let p = self.ptr_mut();
        let mut pde: *mut IDictionaryEnumerator = ptr::null_mut();
        let hr = unsafe {
            (*p).GetEnumerator(&mut pde)
        };
        SUCCEEDED!(hr, DE::from(pde), IDictionary)
    }

    fn remove<K, TOut>(&self, key: K) -> Result<()>
        where K: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt : VARIANT = VARIANT::from(key.into_variant());
        let hr = unsafe {
            (*p).Remove(vt)
        };
        SUCCEEDED!(hr, (), IDictionary)
    }
}

pub trait DictionaryEnumerator where Self: PtrContainer<IDictionaryEnumerator> {
    fn key<DV>(&self) -> Result<DV>
        where DV: From<*mut VARIANT>
    {
        let p = self.ptr_mut();
        let mut vt: VARIANT = unsafe{mem::zeroed()};
        let hr = unsafe {
            (*p).get_key(&mut vt)
        };
        SUCCEEDED!(hr, DV::from(&mut vt), IDictionaryEnumerator)
    }

    fn value<DV>(&self) -> Result<DV>
        where DV: From<*mut VARIANT>
    {
        let p = self.ptr_mut();
        let mut vt: VARIANT = unsafe{mem::zeroed()};
        let hr = unsafe {
            (*p).get_val(&mut vt)
        };
        SUCCEEDED!(hr, DV::from(&mut vt), IDictionaryEnumerator)
    }

    fn entry<DE>(&self) -> Result<DE>
        where DE: From<*mut DictionaryEntry>
    {
        let p = self.ptr_mut();
        let mut de: DictionaryEntry = unsafe {mem::zeroed()};
        let hr = unsafe {
            (*p).get_Entry(&mut de)
        };
        SUCCEEDED!(hr, DE::from(&mut de), IDictionaryEnumerator)
    }
}

pub trait Enumerable where Self: PtrContainer<IEnumerable> {
    fn enumerator<EN>(&self) -> Result<EN>
        where EN: From<*mut IEnumerator>
    {
        let p = self.ptr_mut();
        let mut pie: *mut IEnumerator = ptr::null_mut();
        let hr = unsafe {
            (*p).GetEnumerator(&mut pie)
        };

        SUCCEEDED!(hr, EN::from(pie), IEnumerable)
    }
}

pub trait Enumerator where Self: PtrContainer<IEnumerator> {
    fn move_next(&self) -> Result<bool>{
        let p = self.ptr_mut();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).MoveNext(&mut vb)
        };
        SUCCEEDED!(hr, vb < 0,IEnumerator)
    }
    
    fn current<V>(&self) -> Result<V>
        where V: From<*mut VARIANT>
    {
        let p = self.ptr_mut();
        let mut vt: VARIANT = unsafe{mem::zeroed()};
        let hr = unsafe {
            (*p).get_Current(&mut vt) 
        };
        SUCCEEDED!(hr, V::from(&mut vt),IEnumerator)
    }

    fn reset(&self) -> Result<()>{
        let p = self.ptr_mut();
        let hr = unsafe {
            (*p).Reset()
        };
        SUCCEEDED!(hr, (), IEnumerator)
    }
}

pub trait EqualityComparer where Self: PtrContainer<IEqualityComparer> {
    fn equals<X, Y, TOut, TOut2>(&self, x: X, y: Y) -> Result<bool>
        where X: PtrContainer<TOut>, 
              Y: PtrContainer<TOut2>, 
    {
        let p = self.ptr_mut();
        let xvt : VARIANT = VARIANT::from(x.into_variant());
        let yvt : VARIANT = VARIANT::from(y.into_variant());

        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Equals(xvt, yvt, &mut vb)
        };
        SUCCEEDED!(hr, vb < 0, IEqualityComparer)
    }

    fn hash<V, TOut>(&self, obj: V) -> Result<c_long>
        where V: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt : VARIANT = VARIANT::from(obj.into_variant());
        let mut cl: c_long = 0;
        let hr = unsafe {
            (*p).GetHashCode(vt, &mut cl)
        };
        SUCCEEDED!(hr, cl, IEqualityComparer)
    }
}

pub trait HashCodeProvider where Self: PtrContainer<IHashCodeProvider> {
    fn hash<V, TOut>(&self, obj: V) -> Result<c_long>
        where V: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt : VARIANT = VARIANT::from(obj.into_variant());
        let mut cl: c_long = 0;
        let hr = unsafe {
            (*p).GetHashCode(vt, &mut cl)
        };
        SUCCEEDED!(hr, cl, IHashCodeProvider)
    }
}

pub trait List where Self: PtrContainer<IList> {
    fn item<V>(&self, index: i32) -> Result<V>
        where V: From<*mut VARIANT>
    {
        let p = self.ptr_mut();
        let mut v: VARIANT = unsafe {mem::zeroed()};
        let hr = unsafe {
            (*p).get_Item(index as c_long, &mut v)
        };
        SUCCEEDED!(hr, V::from(&mut v), IList)
    }

    fn item_mut<V, TOut>(&self, index: i32, value: V) -> Result<()>
        where V: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt : VARIANT = VARIANT::from(value.into_variant());
        let hr = unsafe {
            (*p).putref_Item(index, vt)
        };
        SUCCEEDED!(hr, (), IList)
    }

    fn add<V, TOut>(&self, value: V) -> Result<i32>
        where V: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt : VARIANT = VARIANT::from(value.into_variant());
        let mut ret: c_long = 0;
        let hr = unsafe {
            (*p).Add(vt, &mut ret)
        };
        SUCCEEDED!(hr, ret, IList)
    }

    fn contains<V, TOut>(&self, value: V) -> Result<bool>
        where V: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt : VARIANT = VARIANT::from(value.into_variant());
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Contains(vt, &mut vb)
        };
        SUCCEEDED!(hr, vb < 0, IList)
    }

    fn clear(&self) -> Result<()>{
        let p = self.ptr_mut();
        let hr = unsafe {
            (*p).Clear()
        };
        SUCCEEDED!(hr, (), IList)
    }

    fn read_only(&self) -> Result<bool>{
        let p = self.ptr_mut();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).get_IsReadOnly(&mut vb)
        };
        SUCCEEDED!(hr, vb < 0, IList)
    }

    fn fixed_size(&self) -> Result<bool>{
        let p = self.ptr_mut();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).get_IsFixedSize(&mut vb)
        };
        SUCCEEDED!(hr, vb < 0, IList)
    }

    fn index<V, TOut>(&self, value: V) -> Result<i32>
        where V: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt : VARIANT = VARIANT::from(value.into_variant());
        let mut cl: c_long = 0;
        let hr = unsafe {
            (*p).IndexOf(vt, &mut cl)
        };
        SUCCEEDED!(hr, cl, IList)
    }

    fn insert<V, TOut>(&self, index: i32, value: &mut V) -> Result<()>
        where V: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt : VARIANT = VARIANT::from(value.into_variant());
        let hr = unsafe {
            (*p).Insert(index, vt)
        };
        SUCCEEDED!(hr, (), IList)
    } 

    fn remove<V, TOut>(&self, value: &mut V) -> Result<()>
        where V: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt : VARIANT = VARIANT::from(value.into_variant());
        let hr = unsafe {
            (*p).Remove(vt)
        };
        SUCCEEDED!(hr, (), IList)
    }

    fn remove_at(&self, index: i32) -> Result<()>{
        let p = self.ptr_mut();
        let hr = unsafe {
            (*p).RemoveAt(index)
        };
        SUCCEEDED!(hr, (), IList)
    }
}