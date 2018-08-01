//inheritance of IDispatch interfaces:
//$interface.lpVtbl -> *const $interfaceVtbl
//$interfaceVtbl.parent -> IDispatchVtbl
//IDispatchVtbl.parent -> IUnknownVtbl
//IUnknownVtbl {
//  QueryInterface(REFIID, *mut *mut c_void), 
//  AddRef() -> ULONG, 
//  Release() -> ULONG,    
//}
//IDispatchVtbl {
//  GetTypeInfoCount, GetTypeInfo, GetIDsOfNames, Invoke
//}

use std::collections::HashMap;
use std::fmt;
use std::fmt::Display;
use std::ops::Deref;
use std::ptr;

use winapi::ctypes::{c_long};

use winapi::shared::guiddef::{GUID, REFIID, IID_NULL};
use winapi::shared::minwindef::UINT;
use winapi::shared::ntdef::LOCALE_NEUTRAL;
use winapi::shared::wtypes::BSTR;
use winapi::shared::wtypes::VARIANT_BOOL;

use winapi::um::oaidl::{IDispatch};
use winapi::um::oaidl::ITypeInfo;
use winapi::um::oaidl::VARIANT;
use winapi::um::oaidl::SAFEARRAY;
use winapi::um::unknwnbase::{IUnknown};

use mscorlib_sys::system::{_Array, _Attribute, _Version, IComparable, RuntimeTypeHandle};
use mscorlib_sys::system::io::{_FileStream, _Stream};
use mscorlib_sys::system::globalization::_CultureInfo;
use mscorlib_sys::system::reflection::{_Assembly, _AssemblyName, _Binder, _ConstructorInfo, _FieldInfo, _EventInfo, _ManifestResourceInfo, _MemberInfo, 
_MethodBase,_MethodInfo, _Module, _ParameterInfo, _PropertyInfo, _Type, _TypeFilter};
use mscorlib_sys::system::reflection::{BindingFlags, CallingConventions, ICustomAttributeProvider, InterfaceMapping, MemberTypes, MethodAttributes, TypeAttributes};
use mscorlib_sys::system::collections::{DictionaryEntry, ICollection, IComparer, IDictionary, IDictionaryEnumerator, 
IEnumerable, IEnumerator, IEqualityComparer, IHashCodeProvider, IList};
use mscorlib_sys::system::security::policy::_Evidence;


use bstring::{BString};
use variant::{Variant, PhantomDispatch, PhantomUnknown, WrappedDispatch, WrappedUnknown};
use result::{ClrError, SourceLocation, Result};
use safearray::{SafeArray, UnknownSafeArray, DispatchSafeArray};
use struct_wrappers::InterfaceMapping as WrappedInterfaceMapping;

pub trait PtrContainer<T> {
    fn ptr(&self) -> *const T;
    fn ptr_mut(&self) -> *mut T;
    fn from(p: *mut T) -> Self where Self:Sized;
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
        let b = vb > 0;
        SUCCEEDED!(hr, b, ICollection)
    }
}

pub trait Comparable where Self: PtrContainer<IComparable> {
    fn compare<R, T>(&self, rhs: R) -> Result<i32>
        where R: Comparable + PtrContainer<IComparable> + From<*mut IComparable>
    {
        let lhs_ptr: *mut IComparable = self.ptr_mut();
        let rhs_vt = DISPATCH!(rhs: R: IComparable);
        let mut ret: c_long = 0;
        let hr = unsafe {
            (*lhs_ptr).CompareTo(rhs_vt, &mut ret)
        };

        SUCCEEDED!(hr, ret, IComparable)
    }
}

pub trait Comparer where Self: PtrContainer<IComparer> {
    fn compare<L, R, TLeft, TRight>(&self, lhs: L, rhs: R) -> Result<i32>
        where L: PtrContainer<TLeft> + From<*mut TLeft>, 
              R: PtrContainer<TRight> + From<*mut TRight>, 
              TLeft: Deref<Target=IDispatch>, 
              TRight: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let lhs_vt = DISPATCH!(lhs: L: TLeft);
        let rhs_vt = DISPATCH!(rhs: R: TRight);
        let mut ret: c_long = 0;
        let hr = unsafe {
            (*p).Compare(lhs_vt, rhs_vt, &mut ret)
        };
        SUCCEEDED!(hr, ret, IComparer)
    }
}

pub trait Dictionary where Self: PtrContainer<IDictionary> {
    fn item<K, TDispatch, V, TDispatch2>(&self, key: K) -> Result<Variant<V, WrappedUnknown, TDispatch2, PhantomUnknown>>
        where K: PtrContainer<TDispatch> + From<*mut TDispatch>, 
              V: PtrContainer<TDispatch2> + From<*mut TDispatch2>,
              TDispatch: Deref<Target=IDispatch>, 
              TDispatch2: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let vt = DISPATCH!(key: K: TDispatch);
        let p_ret: *mut VARIANT = ptr::null_mut();
        let hr = unsafe {
            (*p).get_Item(vt, p_ret)
        };
        SUCCEEDED!(hr, Variant::from(p_ret), IDictionary)
    }

    fn item_mut<K, V, TDispatch, TDispatch2>(&mut self, key: K, value: V) -> Result<()>
        where K: PtrContainer<TDispatch> + From<*mut TDispatch>, 
              V: PtrContainer<TDispatch2> + From<*mut TDispatch2>, 
              TDispatch: Deref<Target=IDispatch>, 
              TDispatch2: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let kvt = DISPATCH!(key: K: TDispatch);
        let vvt = DISPATCH!(value: V: TDispatch2);
        let hr = unsafe {
            (*p).putref_Item(kvt, vvt)
        };
        SUCCEEDED!(hr, (), IDictionary)
    }

    fn keys<C: Collection>(&self) -> Result<C>
    {
        let p = self.ptr_mut();
        let mut pic: *mut ICollection = ptr::null_mut();
        let hr = unsafe {
            (*p).get_Keys(&mut pic)
        };
        SUCCEEDED!(hr, C::from(pic), IDictionary)
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
        where V: PtrContainer<TOut>, 
              TOut: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let pvt = DISPATCH!(obj: V: TOut);
        let mut pb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Contains(pvt, &mut pb)
        };
        SUCCEEDED!(hr, pb > 0, IDictionary)
    }

    fn add<K, V, TKey, TValue>(&self, key: K, value: V) -> Result<()>
        where K: PtrContainer<TKey>,
              V: PtrContainer<TValue>, 
              TKey: Deref<Target=IDispatch>, 
              TValue: Deref<Target=IDispatch>
    {
        let k = DISPATCH!(key:K:TKey);
        let v = DISPATCH!(value:V:TValue);
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
        SUCCEEDED!(hr, vb > 0, IDictionary)
    }

    fn fixed_size(&self) -> Result<bool>{
        let p = self.ptr_mut();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).get_IsFixedSize(&mut vb)
        };
        SUCCEEDED!(hr, vb > 0, IDictionary)
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
        where K: PtrContainer<TOut> + From<*mut TOut>,
              TOut: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let vt = DISPATCH!(key:K:TOut);
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
        let pvt: *mut VARIANT = ptr::null_mut();
        let hr = unsafe {
            (*p).get_key(pvt)
        };
        SUCCEEDED!(hr, DV::from(pvt), IDictionaryEnumerator)
    }

    fn value<DV>(&self) -> Result<DV>
        where DV: From<*mut VARIANT>
    {
        let p = self.ptr_mut();
        let pvt: *mut VARIANT = ptr::null_mut();
        let hr = unsafe {
            (*p).get_val(pvt)
        };
        SUCCEEDED!(hr, DV::from(pvt), IDictionaryEnumerator)
    }

    fn entry<DE>(&self) -> Result<DE>
        where DE: From<*mut DictionaryEntry>
    {
        let p = self.ptr_mut();
        let pde: *mut DictionaryEntry = ptr::null_mut();
        let hr = unsafe {
            (*p).get_Entry(pde)
        };
        SUCCEEDED!(hr, DE::from(pde), IDictionaryEnumerator)
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
        SUCCEEDED!(hr, vb > 0,IEnumerator)
    }
    
    fn current<V>(&self) -> Result<V>
        where V: From<*mut VARIANT>
    {
        let p = self.ptr_mut();
        let vt: *mut VARIANT = ptr::null_mut();
        let hr = unsafe {
            (*p).get_Current(vt)
        };
        SUCCEEDED!(hr, V::from(vt),IEnumerator)
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
    fn equals<X, Y, TOut>(&self, x: X, y: Y) -> Result<bool>
        where X: PtrContainer<TOut>, 
              Y: PtrContainer<TOut>, 
              TOut: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let xvt = DISPATCH!(x:X:TOut);
        let yvt = DISPATCH!(y:Y:TOut);

        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Equals(xvt, yvt, &mut vb)
        };
        SUCCEEDED!(hr, vb > 0, IEqualityComparer)
    }

    fn hash<V, TOut>(&self, obj: V) -> Result<c_long>
        where V: PtrContainer<TOut>, 
              TOut: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let vt = DISPATCH!(obj:V:TOut);
        let mut cl: c_long = 0;
        let hr = unsafe {
            (*p).GetHashCode(vt, &mut cl)
        };
        SUCCEEDED!(hr, cl, IEqualityComparer)
    }
}

pub trait HashCodeProvider where Self: PtrContainer<IHashCodeProvider> {
    fn hash<V, TOut>(&self, obj: V) -> Result<c_long>
        where V: PtrContainer<TOut> + From<*mut TOut>, 
              TOut: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let vt = DISPATCH!(obj:V:TOut);
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
        let v: *mut VARIANT = ptr::null_mut();
        let hr = unsafe {
            (*p).get_Item(index as c_long, v)
        };
        SUCCEEDED!(hr, V::from(v), IList)
    }

    fn item_mut<V, TOut>(&self, index: i32, value: V) -> Result<()>
        where V: PtrContainer<TOut> +  From<*mut TOut>, 
              TOut: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let vt = DISPATCH!(value:V:TOut);
        let hr = unsafe {
            (*p).putref_Item(index, vt)
        };
        SUCCEEDED!(hr, (), IList)
    }

    fn add<V, TOut>(&self, value: V) -> Result<i32>
        where V: PtrContainer<TOut> +  From<*mut TOut>, 
              TOut: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let vt = DISPATCH!(value:V:TOut);
        let mut ret: c_long = 0;
        let hr = unsafe {
            (*p).Add(vt, &mut ret)
        };
        SUCCEEDED!(hr, ret, IList)
    }

    fn contains<V, TOut>(&self, value: V) -> Result<bool>
        where V: PtrContainer<TOut> +  From<*mut TOut>, 
              TOut: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let vt = DISPATCH!(value:V:TOut);
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Contains(vt, &mut vb)
        };
        SUCCEEDED!(hr, vb > 0, IList)
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
        SUCCEEDED!(hr, vb > 0, IList)
    }

    fn fixed_size(&self) -> Result<bool>{
        let p = self.ptr_mut();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).get_IsFixedSize(&mut vb)
        };
        SUCCEEDED!(hr, vb > 0, IList)
    }

    fn index<V, TOut>(&self, value: V) -> Result<i32>
        where V: PtrContainer<TOut> + From<*mut TOut>, 
              TOut: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let vt = DISPATCH!(value:V:TOut);
        let mut cl: c_long = 0;
        let hr = unsafe {
            (*p).IndexOf(vt, &mut cl)
        };
        SUCCEEDED!(hr, cl, IList)
    }

    fn insert<V, TOut>(&self, index: i32, value: V) -> Result<()>
        where V: PtrContainer<TOut> + From<*mut TOut>,
              TOut: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let vt = DISPATCH!(value:V:TOut);
        let hr = unsafe {
            (*p).Insert(index, vt)
        };
        SUCCEEDED!(hr, (), IList)
    } 

    fn remove<V, TOut>(&self, value: V) -> Result<()>
        where V: PtrContainer<TOut> + From<*mut TOut>, 
              TOut: Deref<Target=IDispatch>
    {
        let p = self.ptr_mut();
        let vt = DISPATCH!(value:V:TOut);
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

pub trait Assembly where Self: PtrContainer<_Assembly> {
    //implemented by Assembly, which also implements IEvidenceFactory, ICustomAttributeProvider, ISerializable
    //Assembly is an abstract class, implemented by AssemblyBuilder (which lives behind _AssemblyBuilder interface)
    /*
    {   
        fn add_ModuleResolve(
            val: *mut _ModuleResolveEventHandler, 
        ) -> HRESULT,
        fn remove_ModuleResolve( 
            val: *mut _ModuleResolveEventHandler,
        ) -> HRESULT,
        fn LoadModule(
            moduleName: BSTR, 
            rawModule: *mut SAFEARRAY,  //byte[]
            pRetVal: *mut *mut _Module,
        ) -> HRESULT,
        fn LoadModule_2( 
            moduleName: BSTR, 
            rawModule: *mut SAFEARRAY, //byte[]
            rawSymbolStore: *mut SAFEARRAY, //byte[]
            pRetVal: *mut *mut _Module,
        ) -> HRESULT,
        fn CreateInstance(
            typeName: BSTR,
            pRetVal: *mut *mut VARIANT,
        ) -> HRESULT,
        fn CreateInstance_2(
            typeName: BSTR, 
            ignoreCase: VARIANT_BOOL, 
            pRetVal: *mut *mut VARIANT,
        ) -> HRESULT,
        fn CreateInstance_3(
            typeName: BSTR, 
            ignoreCase: VARIANT_BOOL, 
            bindingAttr: BindingFlags, 
            Binder: *mut _Binder,
            args: *mut SAFEARRAY, 
            culture: *mut _CultureInfo, 
            activationAttributes: *mut SAFEARRAY, 
            pRetVal: *mut *mut VARIANT,
        ) -> HRESULT,
        
    }*/

    fn global_assembly_cache(&self) -> Result<bool> {
        let p = self.ptr_mut();
        let vb: *mut VARIANT_BOOL = ptr::null_mut();
        let hr = unsafe {
            (*p).get_GlobalAssemblyCache(vb)
        };
        SUCCEEDED!(hr, unsafe{*vb} > 0, _Assembly)
    }

    fn referenced_assemblies<A>(&self) -> Result<DispatchSafeArray<A, _Assembly>> 
        where A: PtrContainer<_Assembly>
    {
        let p = self.ptr_mut();
        let assemblies: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetReferencedAssemblies(assemblies)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*assemblies}), _Assembly)
    }

    fn module<M>(&self, name: String) -> Result<M> 
        where M:  PtrContainer<_Module>
    {
        let p = self.ptr_mut();
        let module: *mut *mut _Module = ptr::null_mut();
        let bs: BString = From::from(name);
        let hr = unsafe {
            (*p).GetModule(bs.as_sys(), module)
        };
        SUCCEEDED!(hr, M::from(unsafe{*module}), _Assembly)
    }

    fn modules<M>(&self, get_resource_modules: Option<bool>) -> Result<UnknownSafeArray<M, _Module>> 
        where M: PtrContainer<_Module>
    {
        let p = self.ptr_mut();
        let modules: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = match get_resource_modules {
            Some(get) => unsafe {
                let vb_mod: VARIANT_BOOL = if get {1} else {0};
                (*p).GetModules_2(vb_mod, modules)
            }, 
            None => unsafe {
                (*p).GetModules(modules)
            }
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*modules}), _Assembly)
    }

    fn loaded_modules<M>(&self, get_resource_modules: Option<bool>) -> Result<UnknownSafeArray<M, _Module>> 
        where M: PtrContainer<_Module> 
    {
        let p = self.ptr_mut();
        let modules: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = match get_resource_modules {
            Some(get) => unsafe {
                let vb_mod: VARIANT_BOOL = if get {1} else {0};
                (*p).GetLoadedModules_2(vb_mod, modules)
            }, 
            None => unsafe {
                (*p).GetLoadedModules(modules)
            }
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*modules}), _Assembly)
    }

    fn create_instance<V>(&self, _type_name: String, _ignore_case: Option<bool>) -> Result<V> 
    {   /*fn CreateInstance(
            typeName: BSTR,
            pRetVal: *mut *mut VARIANT,
        ) -> HRESULT,
        fn CreateInstance_2(
            typeName: BSTR, 
            ignoreCase: VARIANT_BOOL, 
            pRetVal: *mut *mut VARIANT,
        ) -> HRESULT,
        fn CreateInstance_3(
            typeName: BSTR, 
            ignoreCase: VARIANT_BOOL, 
            bindingAttr: BindingFlags, 
            Binder: *mut _Binder,
            args: *mut SAFEARRAY, 
            culture: *mut _CultureInfo, 
            activationAttributes: *mut SAFEARRAY, 
            pRetVal: *mut *mut VARIANT,
        ) -> HRESULT,*/
        unimplemented!();
    }

    fn custom_attributes<T, A>(&self, inherit: bool, attr: Option<T>) -> Result<UnknownSafeArray<A, _Attribute>> 
        where T: PtrContainer<_Type>, 
              A: PtrContainer<_Attribute>
    {
        let p = self.ptr_mut();
        let vb_inherit: VARIANT_BOOL = if inherit {1} else {0};
        let attrs: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = match attr {
            Some(attr) => unsafe {
                let t = attr.ptr_mut();
                (*p).GetCustomAttributes(t, vb_inherit, attrs)
            }, 
            None => unsafe {
                (*p).GetCustomAttributes_2(vb_inherit, attrs)
            }
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*attrs}), _Assembly )
    }
    
    fn manifest_resource_names(&self) -> Result<SafeArray<WrappedDispatch, WrappedUnknown, PhantomDispatch, PhantomUnknown, String>> 
    {
        let p = self.ptr_mut();
        let names: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetManifestResourceNames(names)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*names}), _Assembly )
    }

    fn files<F>(&self, resource_modules: Option<bool>) -> Result<DispatchSafeArray<F, _FileStream>> 
        where F: PtrContainer<_FileStream>
    {
        let p = self.ptr_mut();
        let files: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = match resource_modules {
            Some(get_modules) => unsafe {
                let vb: VARIANT_BOOL = if get_modules {1} else {0};
                (*p).GetFiles_2(vb, files)
            }, 
            None => unsafe {
                (*p).GetFiles(files)
            }
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*files}), _Assembly)
    }

    fn to_str(&self) -> Result<String>{
        let p = self.ptr_mut();
        let bs: *mut BSTR = ptr::null_mut();
        let hr = unsafe {
            (*p).ToString_(bs)
        };
        SUCCEEDED!(hr, BString::from_ptr_safe(unsafe{*bs}).to_string(), _Assembly)
    }

    fn equals<V, TOut>(&self, value: V) -> Result<bool>
        where V: PtrContainer<TOut>, 
              TOut: Deref<Target=IUnknown>
    {
        let p = self.ptr_mut();
        let vt = UNKNOWN!(value:V:TOut);
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Equals(vt, &mut vb)
        };
        SUCCEEDED!(hr, vb > 0, _Assembly)
    }

    fn hashcode(&self) -> Result<i32>{
        let p = self.ptr_mut();
        let mut cl: c_long = 0;
        let hr = unsafe {
            (*p).GetHashCode(&mut cl)
        };
        SUCCEEDED!(hr, cl, _Assembly)
    }

    fn type_of<F>(&self) -> Result<F>
        where F: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let mut t: *mut _Type = ptr::null_mut();
        let hr = unsafe {
            (*p).GetType(&mut t)
        };
        SUCCEEDED!(hr, F::from(t), _Assembly)
    }

    fn codebase(&self) -> Result<String>{
        let p = self.ptr_mut();
        let bs: *mut BSTR = ptr::null_mut();
        let hr = unsafe {
            (*p).get_CodeBase(bs)
        };
        SUCCEEDED!(hr, BString::from_ptr_safe(unsafe{*bs}).to_string(), _Assembly)
    }

    fn escaped_codebase(&self) -> Result<String>{
        let p = self.ptr_mut();
        let bs: *mut BSTR = ptr::null_mut();
        let hr = unsafe {
            (*p).get_EscapedCodeBase(bs)
        };
        SUCCEEDED!(hr, BString::from_ptr_safe(unsafe{*bs}).to_string(), _Assembly)
    }

    fn name<A>(&self) -> Result<A>
        where A: PtrContainer<_AssemblyName>
    {
        let p = self.ptr_mut();
        let an: *mut *mut _AssemblyName = ptr::null_mut();
        let hr = unsafe {
            (*p).GetName(an)
        };
        SUCCEEDED!(hr, A::from(unsafe {*an}), _Assembly)
    }

    fn name_2<A>(&self, use_code_base_after_shadow_copy: bool) -> Result<A>
        where A: PtrContainer<_AssemblyName>
    {
        let p = self.ptr_mut();
        let an: *mut *mut _AssemblyName = ptr::null_mut();
        let hr = unsafe {
            (*p).GetName_2(if use_code_base_after_shadow_copy  {1} else {0} as VARIANT_BOOL, an)
        };
        SUCCEEDED!(hr, A::from(unsafe {*an}), _Assembly)
    }

    fn full_name(&self) -> Result<String>{
        let p = self.ptr_mut();
        let bs: *mut BSTR = ptr::null_mut();
        let hr = unsafe {
            (*p).get_FullName(bs)
        };
        SUCCEEDED!(hr, BString::from_ptr_safe(unsafe{*bs}).to_string(), _Assembly)
    }

    fn entry_point<M>(&self) -> Result<M>
        where M: PtrContainer<_MethodInfo>
    {
        let p = self.ptr_mut();
        let mi: *mut *mut _MethodInfo = ptr::null_mut();
        let hr = unsafe {
            (*p).get_EntryPoint(mi)
        };
        SUCCEEDED!(hr, M::from(unsafe {*mi}),  _Assembly)
    }

    fn type_2<T>(&self, name: &'static str) -> Result<T> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let t: *mut *mut _Type = ptr::null_mut();
        let hr = unsafe {
            (*p).GetType_2(bs.as_sys(), t)
        };
        SUCCEEDED!(hr, T::from(unsafe {*t}),  _Assembly)
    }
    
    fn type_3<T>(&self, name: &'static str, throw_on_error: bool) -> Result<T> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let t: *mut *mut _Type = ptr::null_mut();
        let hr = unsafe {
            (*p).GetType_3(bs.as_sys(), if throw_on_error {1} else {0} as VARIANT_BOOL, t)
        };
        SUCCEEDED!(hr, T::from(unsafe {*t}),  _Assembly)
    }
    
    fn exported_types<S>(&self) -> Result<UnknownSafeArray<S, _Type>> 
        where S: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let sa: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetExportedTypes(sa)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*sa}) , _Assembly)
    }

    fn types<S>(&self) -> Result<UnknownSafeArray<S, _Type>> 
        where S: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let sa: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetTypes(sa)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*sa}), _Assembly)
    }

    fn manifest_resource_stream<T, S>(&self, t: T, name: String) -> Result<S> 
        where T: PtrContainer<_Type>, 
              S: PtrContainer<_Stream>
    {
        let p = self.ptr_mut();
        let t = t.ptr_mut();
        let bs: BString = From::from(name);
        let s: *mut *mut _Stream = ptr::null_mut();
        let hr = unsafe {
            (*p).GetManifestResourceStream(t, bs.as_sys(), s)
        };
        SUCCEEDED!(hr, S::from(unsafe{*s}), _Assembly)
    }

    fn manifest_resource_stream_2<S>(&self, name: String) -> Result<S> 
        where S: PtrContainer<_Stream> 
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let s: *mut *mut _Stream = ptr::null_mut();
        let hr = unsafe {
            (*p).GetManifestResourceStream_2(bs.as_sys(), s)
        };
        SUCCEEDED!(hr, S::from(unsafe {*s}), _Assembly)
    }

    fn file<F>(&self, name: String) -> Result<F> 
        where F: PtrContainer<_FileStream>
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let f: *mut *mut _FileStream = ptr::null_mut();
        let hr = unsafe {
            (*p).GetFile(bs.as_sys(), f)
        };
        SUCCEEDED!(hr, F::from(unsafe {*f}), _Assembly)
    }

    fn manifest_resource_info<I>(&self, name: String) -> Result<I> 
        where I: PtrContainer<_ManifestResourceInfo> 
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let i: *mut *mut _ManifestResourceInfo = ptr::null_mut();
        let hr = unsafe {
            (*p).GetManifestResourceInfo(bs.as_sys(), i)
        };
        SUCCEEDED!(hr, I::from(unsafe {*i}),  _Assembly)
    }
    
    fn location(&self) -> Result<String> {
        let p = self.ptr_mut();
        let bs: *mut BSTR = ptr::null_mut();
        let hr = unsafe {
            (*p).get_Location(bs)
        };
        SUCCEEDED!(hr, BString::from_ptr_safe(unsafe {*bs}).to_string(), _Assembly)
    }

    fn evidence<E>(&self) -> Result<E>
        where E: PtrContainer<_Evidence>
    {
        let p = self.ptr_mut();
        let e: *mut *mut _Evidence = ptr::null_mut();
        let hr = unsafe {
            (*p).get_Evidence(e)
        };
        SUCCEEDED!(hr, E::from(unsafe {*e}),  _Assembly)
    }

    fn is_defined<T>(&self, attr_type: T, inherit: bool) -> Result<bool> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let t = attr_type.ptr_mut();
        let vb_inherit: VARIANT_BOOL = if inherit {1}  else {0};
        let mut vb_ret: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).IsDefined(t, vb_inherit, &mut vb_ret)
        };
        SUCCEEDED!(hr, vb_ret > 0, _Assembly )
    }

    fn type_4<T>(&self, name: String, throw_on_error: bool, ignore_case: bool) -> Result<T> 
        where T: PtrContainer<_Type> 
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let vb_throw: VARIANT_BOOL = if throw_on_error {1} else {0};
        let vb_ignore: VARIANT_BOOL = if ignore_case {1} else {0};
        let t: *mut *mut _Type = ptr::null_mut();
        let hr = unsafe {
            (*p).GetType_4(bs.as_sys(), vb_throw, vb_ignore, t)
        };
        SUCCEEDED!(hr, T::from(unsafe {*t}), _Assembly)
    }

    fn satellite_assembly<A, C, V>(&self, culture: C, version: Option<V>) -> Result<A> 
        where C: PtrContainer<_CultureInfo>, 
              A: PtrContainer<_Assembly>,
              V: PtrContainer<_Version>
    {
        let p = self.ptr_mut();
        let c = culture.ptr_mut();
        let asm: *mut *mut _Assembly = ptr::null_mut();
        let hr = match version {
            Some(version) => {
                let v = version.ptr_mut();
                unsafe {(*p).GetSatelliteAssembly_2(c, v, asm)}
            }, 
            None =>unsafe {
                (*p).GetSatelliteAssembly(c, asm)
            }
        };
        SUCCEEDED!(hr, A::from(unsafe{*asm}), _Assembly)
    }
}

pub trait Type where Self: PtrContainer<_Type> {
    /*
    fn Invoke(
        dispIdMember: DISPID,
        riid: REFIID,
        lcid: LCID,
        wFlags: WORD,
        pDispParams: *mut DISPPARAMS,
        pVarResult: *mut VARIANT,
        pExcepInfo: *mut EXCEPINFO,
        puArgErr: *mut UINT,
    ) -> HRESULT,

    */

    fn constructor(&self) {
        unimplemented!()
    }

    fn properties<PI>(&self, binding_attrs: BindingFlags) -> Result<UnknownSafeArray<PI, _PropertyInfo>> 
        where PI: PtrContainer<_PropertyInfo> 
    {
        let p = self.ptr_mut();
        let psa: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetProperties(binding_attrs, psa)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe{*psa}), _Type)
    }

    fn property<P>(&self, name: String, binding_attrs: BindingFlags) -> Result<P>
        where P: PtrContainer<_PropertyInfo>
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let pppi: *mut *mut _PropertyInfo = ptr::null_mut();
        let hr = unsafe {
            (*p).GetProperty(bs.as_sys(), binding_attrs, pppi)
        };
        SUCCEEDED!(hr, P::from(unsafe{*pppi}), _Type)
    }

    fn fields<F>(&self, binding_attrs: BindingFlags) -> Result<UnknownSafeArray<F, _FieldInfo>> 
        where F: PtrContainer<_FieldInfo> 
    {
        let p = self.ptr_mut();
        let ppsa: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetFields(binding_attrs, ppsa)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*ppsa}), _Type)
    }

    fn field<F>(&self, binding_attrs: BindingFlags) -> Result<UnknownSafeArray<F, _FieldInfo>> 
        where F: PtrContainer<_FieldInfo> 
    {
        let p = self.ptr_mut();
        let ppsa: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetField(binding_attrs, ppsa)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*ppsa}), _Type)
    }

    fn methods<M>(&self, binding_attrs: BindingFlags) -> Result<UnknownSafeArray<M, _MethodInfo>> 
        where M: PtrContainer<_MethodInfo>
    {
        let p = self.ptr_mut();
        let ppsa: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetMethods(binding_attrs, ppsa)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*ppsa}), _Type)
    }

    //still need to implement GetMethod with binder, types, and modifiers
    fn method<M>(&self, name: String, binding_attrs: BindingFlags) -> Result<M> 
        where M: PtrContainer<_MethodInfo> 
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let ppm: *mut *mut _MethodInfo = ptr::null_mut();
        let hr = unsafe {
            (*p).GetMethod_2(bs.as_sys(), binding_attrs, ppm)
        };
        SUCCEEDED!(hr, M::from(unsafe{*ppm}), _Type)
    }
    fn interface_map<T, P, P2, M>(&self, interface_type: T) -> Result<WrappedInterfaceMapping<P, P2, M>>
        where T: PtrContainer<_Type>, 
              P: PtrContainer<_Type>, 
              P2: PtrContainer<_Type>, 
              M: PtrContainer<_MethodInfo>
    {
        let p = self.ptr_mut();
        let t = interface_type.ptr_mut();
        let ppim: *mut *mut InterfaceMapping = ptr::null_mut();
        let hr = unsafe {
            (*p).GetInterfaceMap(t, ppim)
        };
        SUCCEEDED!(hr, WrappedInterfaceMapping::from(unsafe{**ppim}), _Type)
    }

    fn instance_of_type<VD, VU, TDispatch, TUnknown>(&self, variant: Variant<VD, VU, TDispatch, TUnknown>) -> Result<bool> 
        where VD: PtrContainer<TDispatch>, 
              TDispatch: Deref<Target=IDispatch>, 
              VU: PtrContainer<TUnknown>, 
              TUnknown: Deref<Target=IUnknown> 
    {
        let p = self.ptr_mut();
        let v = VARIANT::from(variant);
        let vb: *mut VARIANT_BOOL = ptr::null_mut();
        let hr = unsafe {
            (*p).IsInstanceOfType(v, vb)
        };
        SUCCEEDED!(hr, unsafe{*vb} > 0, _Type)        
    }

    fn assignable_from<T>(&self, test_type: T) -> Result<bool> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let t = test_type.ptr_mut();
        let vb: *mut VARIANT_BOOL = ptr::null_mut();
        let hr = unsafe {
            (*p).IsAssignableFrom(t, vb)
        };
        SUCCEEDED!(hr, unsafe{*vb} > 0, _Type)        
    }

    fn subclass_of<T>(&self, test_type: T) -> Result<bool> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let t = test_type.ptr_mut();
        let vb: *mut VARIANT_BOOL = ptr::null_mut();
        let hr = unsafe {
            (*p).IsSubclassOf(t, vb)
        };
        SUCCEEDED!(hr, unsafe{*vb} > 0, _Type)        
    }

    fn find_members(&self)
    {
        unimplemented!()
    }

    fn default_members<M>(&self) -> Result<UnknownSafeArray<M, _MemberInfo>>
        where M: PtrContainer<_MemberInfo>
    {
        let p = self.ptr_mut();
        let ppm: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetDefaultMembers(ppm)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe{*ppm}), _Type)
    }

    fn invoke_member(&self) {
        unimplemented!()
    }

    fn members<M>(&self, binding_attr: BindingFlags) -> Result<UnknownSafeArray<M, _MemberInfo>> 
        where M: PtrContainer<_MemberInfo>
    {
        let p = self.ptr_mut();
        let ppsa: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetMembers(binding_attr, ppsa)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*ppsa}), _Type)
    }

    fn member<M>(&self, name: String, member_types: Option<MemberTypes>, binding_flags: BindingFlags) -> Result<UnknownSafeArray<M, _MemberInfo>>
        where M: PtrContainer<_MemberInfo>
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let ppm: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = match member_types {
            Some(member_types) => unsafe {
                (*p).GetMember(bs.as_sys(), member_types, binding_flags, ppm)
            }, 
            None => unsafe {
                (*p).GetMember_2(bs.as_sys(), binding_flags, ppm)
            }
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe{*ppm}), _Type)
    }

    fn nested_type<T>(&self, name: String, binding_flags: BindingFlags) -> Result<T> 
        where T: PtrContainer<_Type> 
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let ppt: *mut *mut _Type = ptr::null_mut();
        let hr = unsafe {
            (*p).GetNestedType(bs.as_sys(), binding_flags, ppt)
        };
        SUCCEEDED!(hr, T::from(unsafe {*ppt}), _Type)
    }

    fn nested_types<T>(&self, binding_flags: BindingFlags) -> Result<UnknownSafeArray<T, _Type>> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let psa: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetNestedTypes(binding_flags, psa)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*psa}), _Type)
    }

    fn events<E>(&self, binding_flags: Option<BindingFlags>) -> Result<UnknownSafeArray<E, _EventInfo>> 
        where E: PtrContainer<_EventInfo>
    {
        let p = self.ptr_mut();
        let e: *mut *mut SAFEARRAY = ptr::null_mut();

        let hr = match binding_flags {
            Some(flags) => unsafe {
                (*p).GetEvents_2(flags, e)
            }, 
            None => unsafe {
                (*p).GetEvents(e)
            }
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe{*e}), _Type)
    }

    fn event<E>(&self, name: String, flags: BindingFlags) -> Result<E>
        where E: PtrContainer<_EventInfo> 
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let e: *mut *mut _EventInfo = ptr::null_mut();
        let hr = unsafe {
            (*p).GetEvent(bs.as_sys(), flags, e)
        };
        SUCCEEDED!(hr, E::from(unsafe {*e}), _Type)
    }

    fn find_interfaces<TF, T, V, E>(&self, _filter: TF, _criteria: V) -> Result<UnknownSafeArray<T, _Type>> 
        where TF: PtrContainer<_TypeFilter>,
              T: PtrContainer<_Type>,
              V: PtrContainer<E>
    {
        unimplemented!();
    }
    fn interfaces<T>(&self) -> Result<UnknownSafeArray<T, _Type>> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let psa: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetInterfaces(psa)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*psa}), _Type)
    }

    fn interface<T>(&self, name: String, ignore_case: bool) -> Result<T> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let t: *mut *mut _Type = ptr::null_mut();
        let vb: VARIANT_BOOL = if ignore_case{1} else{0};
        let hr = unsafe {
            (*p).GetInterface(bs.as_sys(), vb, t)
        };
        SUCCEEDED!(hr, T::from(unsafe{*t}), _Type)
    }

    fn constructors<C>(&self, binding_attrs: BindingFlags) -> Result<UnknownSafeArray<C, _ConstructorInfo>> 
        where C: PtrContainer<_ConstructorInfo>
    {
        let p = self.ptr_mut();
        let psa: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetConstructors(binding_attrs, psa)
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe{*psa}), _Type)
    }

    fn defined<T>(&self, attr_type: T, inherit: bool) -> Result<bool>
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let vb: VARIANT_BOOL = if inherit{1} else {0};
        let attr = attr_type.ptr_mut();
        let ret: *mut VARIANT_BOOL = ptr::null_mut();
        let hr = unsafe {
            (*p).IsDefined(attr, vb, ret)
        };
        SUCCEEDED!(hr, unsafe{*ret} > 0, _Type)
    }

    fn custom_attributes<T, A>(&self, inherit: bool, attr_type: Option<T>) -> Result<UnknownSafeArray<A, _Attribute>>
        where A: PtrContainer<_Attribute>, 
              T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let vb: VARIANT_BOOL = if inherit {1} else {0};
        let psa: *mut *mut SAFEARRAY = ptr::null_mut();
        let hr = match attr_type {
            Some(attr) => unsafe {
                let t = attr.ptr_mut();
                (*p).GetCustomAttributes(t, vb, psa)
            },
            None => unsafe {
                (*p).GetCustomAttributes_2(vb, psa)
            }
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*psa}), _Type)
    }

    PROPERTY!{get_DeclaringType _Type { get { declaring_type(_Type) }}}
    PROPERTY!{get_ReflectedType _Type { get { reflected_type(_Type) }}}
    PROPERTY!{get_Guid _Type { get { guid_of(GUID) }}}
    PROPERTY!{get_Module _Type{ get { module(_Module) }}}
    PROPERTY!{get_Assembly _Type{ get { assembly(_Assembly) }}}
    PROPERTY!{get_TypeHandle _Type{ get { runtime_type_handle(RuntimeTypeHandle) }}}
    PROPERTY!{get_BaseType _Type{ get { base_type(_Type) }}}
    PROPERTY!{GetElementType _Type{ get { element_type(_Type) }}}
    PROPERTY!{get_Attributes _Type{ get { attrs(TypeAttributes) }}}
    PROPERTY!{get_IsNotPublic _Type{ get { not_public(VARIANT_BOOL) }}}
    PROPERTY!{get_IsPublic _Type{ get { public(VARIANT_BOOL) }}}
    PROPERTY!{get_IsNestedPublic _Type{ get { nested_public(VARIANT_BOOL) }}}
    PROPERTY!{get_IsNestedPrivate _Type{ get { nested_private(VARIANT_BOOL) }}}
    PROPERTY!{get_IsNestedFamily _Type{ get { nested_family(VARIANT_BOOL) }}}
    PROPERTY!{get_IsNestedAssembly _Type{ get { nested_assembly(VARIANT_BOOL) }}}
    PROPERTY!{get_IsNestedFamANDAssem _Type{ get { nested_fam_and_assem(VARIANT_BOOL) }}}
    PROPERTY!{get_IsNestedFamORAssem _Type{ get { nested_fam_or_assem(VARIANT_BOOL) }}}
    PROPERTY!{get_IsAutoLayout _Type{ get { auto_layout(VARIANT_BOOL) }}}
    PROPERTY!{get_IsLayoutSequential _Type{ get { layout_sequential(VARIANT_BOOL) }}}
    PROPERTY!{get_IsExplicitLayout _Type{ get { explicit_layout(VARIANT_BOOL) }}}
    PROPERTY!{get_IsClass _Type{ get { is_class(VARIANT_BOOL) }}}
    PROPERTY!{get_IsInterface _Type{ get { is_interface(VARIANT_BOOL) }}}
    PROPERTY!{get_IsValueType _Type{ get { value_type(VARIANT_BOOL) }}}
    PROPERTY!{get_IsAbstract _Type{ get { abstract_(VARIANT_BOOL) }}}
    PROPERTY!{get_IsSealed _Type{ get { sealed(VARIANT_BOOL) }}}
    PROPERTY!{get_IsEnum _Type{ get { enum_(VARIANT_BOOL) }}}
    PROPERTY!{get_IsSpecialName _Type{ get { special_name(VARIANT_BOOL) }}}
    PROPERTY!{get_IsImport _Type{ get { import(VARIANT_BOOL) }}}
    PROPERTY!{get_IsSerializable _Type{ get { serializable(VARIANT_BOOL) }}}
    PROPERTY!{get_IsAnsiClass _Type{ get { ansi_class(VARIANT_BOOL) }}}
    PROPERTY!{get_IsUnicodeClass _Type{ get { unicode_class(VARIANT_BOOL) }}}
    PROPERTY!{get_IsAutoClass _Type{ get { auto_class(VARIANT_BOOL) }}}
    PROPERTY!{get_IsArray _Type{ get { array(VARIANT_BOOL) }}}
    PROPERTY!{get_IsByRef _Type{ get { by_ref(VARIANT_BOOL) }}}
    PROPERTY!{get_IsPointer _Type{ get { is_ptr(VARIANT_BOOL) }}}
    PROPERTY!{get_IsPrimitive _Type{ get { primitive(VARIANT_BOOL) }}}
    PROPERTY!{get_IsCOMObject _Type{ get { com_obj(VARIANT_BOOL) }}}
    PROPERTY!{get_HasElementType _Type{ get { has_element_type(VARIANT_BOOL) }}}
    PROPERTY!{get_IsContextful _Type{ get { contextful(VARIANT_BOOL) }}}
    PROPERTY!{get_IsMarshalByRef _Type{ get { marshal_by_ref(VARIANT_BOOL) }}}
    PROPERTY!{get_FullName _Type{ get {full_name(u16)}}}
    PROPERTY!{get_Namespace _Type{ get {namespace(u16)}}}
    PROPERTY!{get_AssemblyQualifiedName _Type{ get {assembly_qual_name(u16)}}}
    PROPERTY!{GetArrayRank _Type{ get {array_rank(c_long)}}}
    PROPERTY!{get_UnderlyingSystemType _Type { get {underlying_system_type(_Type)}}}

    fn name(&self) -> Result<String> {
        let p = self.ptr_mut();
        let pbs: *mut BSTR = ptr::null_mut();
        let hr = unsafe {
            (*p).get_Name(pbs)
        };
        
        SUCCEEDED!(hr, BString::from_ptr_safe(unsafe {*pbs}).to_string(), _Type)
    }

    fn member_types(&self) -> Result<MemberTypes>{
        let p = self.ptr_mut();
        let mut mt: MemberTypes = 0;
        let hr = unsafe {
            (*p).get_MemberType(&mut mt)
        };
        SUCCEEDED!(hr, mt, _Type)
    }

    fn equals<V, TOut>(&self, value: V) -> Result<bool>
        where V: PtrContainer<TOut>, 
              TOut: Deref<Target=IUnknown>
    {
        let p = self.ptr_mut();
        let vt = UNKNOWN!(value:V:TOut);
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Equals(vt, &mut vb)
        };
        SUCCEEDED!(hr, vb > 0, _Type)
    }

    fn equals_2<T>(&self, obj: T) -> Result<bool> 
        where T: PtrContainer<_Type> 
    {
        let p = self.ptr_mut();
        let t = obj.ptr_mut();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Equals_2(t, &mut vb)
        };
        SUCCEEDED!(hr, vb > 0, _Type)
    }

    fn hashcode(&self) -> Result<i32>{
        let p = self.ptr_mut();
        let mut cl: c_long = 0;
        let hr = unsafe {
            (*p).GetHashCode(&mut cl)
        };
        SUCCEEDED!(hr, cl, _Type)
    }

    fn type_of<F>(&self) -> Result<F>
        where F: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let mut t: *mut _Type = ptr::null_mut();
        let hr = unsafe {
            (*p).GetType(&mut t)
        };
        SUCCEEDED!(hr, F::from(t), _Type)
    }

    fn invoke<T>(&self, _name: Option<&str>, _id: Option<i32>, _args: &[&dyn PtrContainer<T>]) {
        unimplemented!();
    }

    fn ids_of_names(&self, names: Vec<&'static str>) -> Result<HashMap<String, i32>> {
        let p = self.ptr_mut();
        let riid: REFIID = &IID_NULL;
        let c_names = names.len();
        let mut disp_ids: Vec<i32> = Vec::with_capacity(c_names);
        let copy_names = names.clone();
        let mut us_names: Vec<*mut u16> = names.into_iter().map(|name| unsafe {
            let bs: BString = From::from(name);
            bs.as_sys()
        }).collect();
        let hr = unsafe {
            (*p).GetIDsOfNames(riid, us_names[..].as_mut_ptr(), c_names as UINT, LOCALE_NEUTRAL, disp_ids[..].as_mut_ptr())
        };
        SUCCEEDED!(hr, {
            let mut hm = HashMap::new();
            copy_names.into_iter().zip(disp_ids).for_each(|(x, y)| {
                hm.insert(String::from(x), y);
            });
            hm
        }, _Type)
    }

    fn type_info<T>(&self, index: u8) -> Result<T>
        where T: PtrContainer<ITypeInfo>
    {
        let p = self.ptr_mut();
        let iti: *mut *mut ITypeInfo = ptr::null_mut();
        let hr = unsafe {
            (*p).GetTypeInfo(index as UINT, LOCALE_NEUTRAL,iti )
        };
        SUCCEEDED!(hr, T::from(unsafe {*iti}), _Type)
    }

    fn type_info_count(&self) -> Result<u32> {
        let p = self.ptr_mut();
        let count: *mut UINT = ptr::null_mut();
        let hr = unsafe {
            (*p).GetTypeInfoCount(count)
        };
        SUCCEEDED!(hr, unsafe { *count }, _Type)
    }

    fn to_str(&self) -> Result<BString>{
        let p = self.ptr_mut();
        let mut bs: BSTR = ptr::null_mut();
        let hr = unsafe {
            (*p).ToString_(&mut bs)
        };
        SUCCEEDED!(hr, BString::from_ptr_safe(bs), _Type)
    }
}

pub trait MemberInfo where Self: PtrContainer<_MemberInfo> {
    fn to_str(&self) -> Result<String> {
        let p = self.ptr_mut();
        let pbs: *mut BSTR = ptr::null_mut();
        let hr = unsafe {
            (*p).ToString_(pbs)
        };
        SUCCEEDED!(hr, BString::from_ptr_safe(unsafe{*pbs}).to_string(), _MemberInfo )
    }

    fn equals<V, TOut>(&self, value: V) -> Result<bool>
        where V: PtrContainer<TOut>, 
              TOut: Deref<Target=IUnknown>
    {
        let p = self.ptr_mut();
        let vt = UNKNOWN!(value:V:TOut);
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Equals(vt, &mut vb)
        };
        SUCCEEDED!(hr, vb > 0, _MemberInfo)
    }

    fn hashcode(&self) -> Result<i32>{
        let p = self.ptr_mut();
        let mut cl: c_long = 0;
        let hr = unsafe {
            (*p).GetHashCode(&mut cl)
        };
        SUCCEEDED!(hr, cl, _MemberInfo)
    }

    fn type_of<F>(&self) -> Result<F>
        where F: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let mut t: *mut _Type = ptr::null_mut();
        let hr = unsafe {
            (*p).GetType(&mut t)
        };
        SUCCEEDED!(hr, F::from(t), _MemberInfo)
    }

    fn member_types(&self) -> Result<MemberTypes>{
        let p = self.ptr_mut();
        let mut mt: MemberTypes = 0;
        let hr = unsafe {
            (*p).get_MemberType(&mut mt)
        };
        SUCCEEDED!(hr, mt, _Type)
    }

    PROPERTY!{get_Name _MemberInfo { get {name(u16)}}}
    PROPERTY!{get_DeclaringType _MemberInfo { get {declaring_type(_Type)}}}
    PROPERTY!{get_ReflectedType _MemberInfo { get {reflected_type(_Type)}}}
    
    fn custom_attributes<T>(&self, inherit: bool, attr_type: Option<T>) -> Result<UnknownSafeArray<T, _Type>>
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let ppsa: *mut *mut SAFEARRAY = ptr::null_mut();
        let vb: VARIANT_BOOL = if inherit {1} else {0};
        let hr = match attr_type {
            Some(attr_type) => unsafe {
                let t = attr_type.ptr_mut();
                (*p).GetCustomAttributes(t, vb, ppsa)
            }, 
            None => unsafe {
                (*p).GetCustomAttributes_2(vb, ppsa)
            }
        };
        SUCCEEDED!(hr, SafeArray::from(unsafe {*ppsa}), _MemberInfo)
    }

    fn is_defined<T>(&self, attr_type: T, inherit: bool) -> Result<bool> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let vb: VARIANT_BOOL = if inherit {1} else {0};
        let ret: *mut VARIANT_BOOL = ptr::null_mut();
        let t = attr_type.ptr_mut();
        let hr = unsafe {
            (*p).IsDefined(t, vb, ret)
        };
        SUCCEEDED!(hr, unsafe{*ret} > 0, _MemberInfo)
    }
}

pub trait MethodBase where Self: PtrContainer<_MethodBase> {
    PROPERTY!{get_IsPublic _MethodBase { get {public(VARIANT_BOOL)}}}
    PROPERTY!{get_IsPrivate _MethodBase { get {private(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFamily _MethodBase { get {family(VARIANT_BOOL)}}}
    PROPERTY!{get_IsAssembly _MethodBase { get {assem(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFamilyAndAssembly _MethodBase { get {family_and_assembly(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFamilyOrAssembly _MethodBase { get {family_or_assembly(VARIANT_BOOL)}}}
    PROPERTY!{get_IsStatic _MethodBase { get {is_static(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFinal _MethodBase { get {is_final(VARIANT_BOOL)}}}
    PROPERTY!{get_IsHideBySig _MethodBase { get {hide_by_sig(VARIANT_BOOL)}}}
    PROPERTY!{get_IsAbstract _MethodBase { get {is_abstract(VARIANT_BOOL)}}}
    PROPERTY!{get_IsSpecialName _MethodBase { get {special_name(VARIANT_BOOL)}}}
    PROPERTY!{get_IsConstructor _MethodBase { get {is_constructor(VARIANT_BOOL)}}}
}

pub trait MethodInfo where Self: PtrContainer<_MethodInfo> {
    PROPERTY!{get_Attributes _MethodInfo { get {attributes(MethodAttributes)}}}
    PROPERTY!{get_CallingConvention _MethodInfo { get {calling_convention(CallingConventions)}}}
    
    PROPERTY!{get_IsPublic _MethodInfo { get {public(VARIANT_BOOL)}}}
    PROPERTY!{get_IsPrivate _MethodInfo { get {private(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFamily _MethodInfo { get {family(VARIANT_BOOL)}}}
    PROPERTY!{get_IsAssembly _MethodInfo { get {assem(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFamilyAndAssembly _MethodInfo { get {family_and_assembly(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFamilyOrAssembly _MethodInfo { get {family_or_assembly(VARIANT_BOOL)}}}
    PROPERTY!{get_IsStatic _MethodInfo { get {is_static(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFinal _MethodInfo { get {is_final(VARIANT_BOOL)}}}
    PROPERTY!{get_IsHideBySig _MethodInfo { get {hide_by_sig(VARIANT_BOOL)}}}
    PROPERTY!{get_IsAbstract _MethodInfo { get {is_abstract(VARIANT_BOOL)}}}
    PROPERTY!{get_IsSpecialName _MethodInfo { get {special_name(VARIANT_BOOL)}}}
    PROPERTY!{get_IsConstructor _MethodInfo { get {is_constructor(VARIANT_BOOL)}}}
    PROPERTY!{get_returnType _MethodInfo { get {return_type(_Type)}}}
    PROPERTY!{get_ReturnTypeCustomAttributes _MethodInfo { get {return_type_custom_attrs(ICustomAttributeProvider)}}}
    PROPERTY!{GetBaseDefinition _MethodInfo { get {base_definition(_MethodInfo)}}}
}

pub trait ConstructorInfo where Self: PtrContainer<_ConstructorInfo>
{
    PROPERTY!{GetType _ConstructorInfo { get {type_of(_Type)}}}
    PROPERTY!{get_name _ConstructorInfo { get {name(u16)}}}
    PROPERTY!{get_DeclaringType _ConstructorInfo { get {declaring_type(_Type)}}}
    PROPERTY!{get_ReflectedType _ConstructorInfo { get {reflected_type(_Type)}}}
    PROPERTY!{get_IsPublic _ConstructorInfo { get {public(VARIANT_BOOL)}}}
    PROPERTY!{get_IsPrivate _ConstructorInfo { get {private(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFamily _ConstructorInfo { get {family(VARIANT_BOOL)}}}
    PROPERTY!{get_IsAssembly _ConstructorInfo { get {assem(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFamilyAndAssembly _ConstructorInfo { get {family_and_assembly(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFamilyOrAssembly _ConstructorInfo { get {family_or_assembly(VARIANT_BOOL)}}}
    PROPERTY!{get_IsStatic _ConstructorInfo { get {is_static(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFinal _ConstructorInfo { get {is_final(VARIANT_BOOL)}}}
    PROPERTY!{get_IsHideBySig _ConstructorInfo { get {hide_by_sig(VARIANT_BOOL)}}}
    PROPERTY!{get_IsAbstract _ConstructorInfo { get {is_abstract(VARIANT_BOOL)}}}
    PROPERTY!{get_IsSpecialName _ConstructorInfo { get {special_name(VARIANT_BOOL)}}}
    PROPERTY!{get_IsConstructor _ConstructorInfo { get {is_constructor(VARIANT_BOOL)}}}   
}

pub trait FieldInfo where Self: PtrContainer<_FieldInfo>  {
    PROPERTY!{get_ToString _FieldInfo { get {to_str(u16)}}}
    PROPERTY!{GetType _FieldInfo { get {type_of(_Type)}}}
    PROPERTY!{get_name _FieldInfo { get {name(u16)}}}
    PROPERTY!{get_DeclaringType _FieldInfo { get {declaring_type(_Type)}}}
    PROPERTY!{get_ReflectedType _FieldInfo { get {reflected_type(_Type)}}}
    PROPERTY!{get_FieldType _FieldInfo { get {field_type(_Type)}}}
    PROPERTY!{get_IsPublic _FieldInfo { get {public(VARIANT_BOOL)}}}
    PROPERTY!{get_IsPrivate _FieldInfo { get {private(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFamily _FieldInfo { get {family(VARIANT_BOOL)}}}
    PROPERTY!{get_IsAssembly _FieldInfo { get {assem(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFamilyAndAssembly _FieldInfo { get {family_and_assembly(VARIANT_BOOL)}}}
    PROPERTY!{get_IsFamilyOrAssembly _FieldInfo { get {family_or_assembly(VARIANT_BOOL)}}}
    PROPERTY!{get_IsStatic _FieldInfo { get {is_static(VARIANT_BOOL)}}}
    PROPERTY!{get_IsInitOnly _FieldInfo { get {init_only(VARIANT_BOOL)}}}
    PROPERTY!{get_IsLiteral _FieldInfo { get {literal(VARIANT_BOOL)}}}
    PROPERTY!{get_IsNotSerialized _FieldInfo { get {not_serialized(VARIANT_BOOL)}}}
    PROPERTY!{get_IsSpecialName _FieldInfo { get {special_name(VARIANT_BOOL)}}}
    PROPERTY!{get_IsPinvokeImpl _FieldInfo { get {pinvoke_impl(VARIANT_BOOL)}}}
}

pub trait PropertyInfo where Self: PtrContainer<_PropertyInfo> {
    PROPERTY!{get_ToString _PropertyInfo { get {to_str(u16)}}}
    PROPERTY!{GetType _PropertyInfo { get {type_of(_Type)}}}
    PROPERTY!{get_name _PropertyInfo { get {name(u16)}}}

    PROPERTY!{get_DeclaringType _PropertyInfo { get {declaring_type(_Type)}}}
    PROPERTY!{get_ReflectedType _PropertyInfo { get {reflected_type(_Type)}}}
    PROPERTY!{get_PropertyType _PropertyInfo { get {property_type(_Type)}}}
    //PROPERTY!{get_Attributes _FieldInfo { get {public(VARIANT_BOOL)}}}
    PROPERTY!{get_CanRead _FieldInfo { get {can_read(VARIANT_BOOL)}}}
    PROPERTY!{get_CanWrite _FieldInfo { get {can_write(VARIANT_BOOL)}}}
    PROPERTY!{GetGetMethod_2 _FieldInfo { get {getter(_MethodInfo)}}}
    PROPERTY!{GetSetMethod_2 _FieldInfo { get {setter(_MethodInfo)}}}
    PROPERTY!{get_IsSpecialName _FieldInfo { get {special_name(VARIANT_BOOL)}}}
}

pub trait EventInfo where Self: PtrContainer<_EventInfo> {
    PROPERTY!{get_ToString _EventInfo { get {to_str(u16)}}}
    PROPERTY!{GetType _EventInfo { get {type_of(_Type)}}}
    PROPERTY!{get_name _EventInfo { get {name(u16)}}}

    PROPERTY!{get_DeclaringType _EventInfo { get {declaring_type(_Type)}}}
    PROPERTY!{get_ReflectedType _EventInfo { get {reflected_type(_Type)}}}
    PROPERTY!{get_IsMulticast _EventInfo { get {multicast(VARIANT_BOOL)}}}
    PROPERTY!{get_EventHandlerType _EventInfo { get {event_handler_type(_Type)}}}
}

pub trait ParameterInfo where Self: PtrContainer<_ParameterInfo> {

}

pub trait Module where Self: PtrContainer<_Module> {
    
}

pub trait AssemblyName where Self: PtrContainer<_AssemblyName> {

}

pub trait Binder where Self: PtrContainer<_Binder> {

}

pub struct ClrType {
    ptr: *mut _Type,
}

impl PtrContainer<_Type> for ClrType {
    fn ptr(&self) -> *const _Type {
        self.ptr
    }
    fn ptr_mut(&self) -> *mut _Type {
        self.ptr
    }
    
    fn from(pt: *mut _Type) -> ClrType {
        ClrType{ ptr: pt}
    }
}

macro_rules! BLANKET_IMPLS {
    ($({$tr:ty, $ptr_ty:ty},)*) => {
        $(
            impl<T:PtrContainer<$ptr_ty>> $tr for T{}
        )*
    };
}

BLANKET_IMPLS!{
    {Collection, ICollection},
    {Comparable, IComparable},
    {Comparer, IComparer},
    {Dictionary, IDictionary},
    {DictionaryEnumerator, IDictionaryEnumerator},
    {Enumerable, IEnumerable}, 
    {Enumerator, IEnumerator},
    {EqualityComparer, IEqualityComparer}, 
    {HashCodeProvider, IHashCodeProvider}, 
    {List, IList}, 
    {Assembly, _Assembly}, 
    {Type, _Type}, 
    {MemberInfo, _MemberInfo}, 
    {MethodBase, _MethodBase}, 
    {ConstructorInfo, _ConstructorInfo}, 
    {FieldInfo, _FieldInfo}, 
    {PropertyInfo, _PropertyInfo}, 
    {EventInfo, _EventInfo}, 
    {ParameterInfo, _ParameterInfo}, 
    {Module, _Module}, 
    {AssemblyName, _AssemblyName},
    {Binder, _Binder},
}

impl Display for ClrType {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        let s = self.to_str().unwrap_or(From::from("ClrType"));
        write!(f, "{:?}", s)
    }
}



