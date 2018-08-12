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
use std::mem;
use std::ptr;

use winapi::ctypes::{c_long};

use winapi::shared::guiddef::{GUID, REFIID, IID_NULL};
use winapi::shared::minwindef::UINT;
use winapi::shared::ntdef::LOCALE_NEUTRAL;
use winapi::shared::wtypes::BSTR;
use winapi::shared::wtypes::VARIANT_BOOL;

use winapi::um::oaidl::IDispatch;
use winapi::um::oaidl::ITypeInfo;
use winapi::um::oaidl::SAFEARRAY;
use winapi::um::unknwnbase::{IUnknown};

use mscorlib_sys::system::{_Attribute, _Version, RuntimeTypeHandle};
use mscorlib_sys::system::io::{_FileStream, _Stream};
use mscorlib_sys::system::globalization::_CultureInfo;
use mscorlib_sys::system::reflection::{_Assembly, _AssemblyName, _Binder, _ConstructorInfo, _FieldInfo, _EventInfo, _ManifestResourceInfo, _MemberInfo, 
_MethodBase,_MethodInfo, _Module, _ParameterInfo, _PropertyInfo, _Type, _TypeFilter};
use mscorlib_sys::system::reflection::{BindingFlags, CallingConventions, ICustomAttributeProvider, InterfaceMapping, MemberTypes, MethodAttributes, TypeAttributes};

use mscorlib_sys::system::security::policy::_Evidence;


use bstring::{BString};

use new_safearray::RSafeArray;
use new_variant::Variant;
use result::{ClrError, SourceLocation, Result};
use struct_wrappers::InterfaceMapping as WrappedInterfaceMapping;

pub trait PtrContainer<T> {
    fn ptr(&self) -> *const T;
    fn ptr_mut(&self) -> *mut T;
    fn from(p: *mut T) -> Self where Self:Sized;
    fn into_variant(&self) -> Variant;
}

//#[incomplete]
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
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).get_GlobalAssemblyCache(&mut vb)
        };
        SUCCEEDED!(hr, vb < 0, _Assembly)
    }

    fn referenced_assemblies<A>(&self) -> Result<Vec<A>> 
        where A: PtrContainer<_Assembly>
    {
        let p = self.ptr_mut();
        let mut passemblies: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetReferencedAssemblies(&mut passemblies)
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Dispatch, passemblies, *mut IDispatch, *mut _Assembly, A}, _Assembly)
    }

    fn module<M>(&self, name: String) -> Result<M> 
        where M:  PtrContainer<_Module>
    {
        let p = self.ptr_mut();
        let mut pmodule: *mut _Module = ptr::null_mut();
        let bs: BString = From::from(name);
        let hr = unsafe {
            (*p).GetModule(bs.as_sys(), &mut pmodule)
        };
        SUCCEEDED!(hr, M::from(pmodule), _Assembly)
    }

    fn modules<M>(&self, get_resource_modules: Option<bool>) -> Result<Vec<M>> 
        where M: PtrContainer<_Module>
    {
        let p = self.ptr_mut();
        let mut pmodules: *mut SAFEARRAY = ptr::null_mut();
        let hr = match get_resource_modules {
            Some(get) => unsafe {
                let vb_mod: VARIANT_BOOL = if get {-1} else {0};
                (*p).GetModules_2(vb_mod, &mut pmodules)
            }, 
            None => unsafe {
                (*p).GetModules(&mut pmodules)
            }
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, pmodules, *mut IUnknown, *mut _Module, M}, _Assembly)
    }

    fn loaded_modules<M>(&self, get_resource_modules: Option<bool>) -> Result<Vec<M>> 
        where M: PtrContainer<_Module> 
    {
        let p = self.ptr_mut();
        let mut pmodules: *mut SAFEARRAY = ptr::null_mut();
        let hr = match get_resource_modules {
            Some(get) => unsafe {
                let vb_mod: VARIANT_BOOL = if get {-1} else {0};
                (*p).GetLoadedModules_2(vb_mod, &mut pmodules)
            }, 
            None => unsafe {
                (*p).GetLoadedModules(&mut pmodules)
            }
        };
        
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, pmodules, *mut IUnknown, *mut _Module, M}, _Assembly)
    }
    //#[incomplete]
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

    fn custom_attributes<T, A>(&self, inherit: bool, attr: Option<T>) -> Result<Vec<A>> 
        where T: PtrContainer<_Type>, 
              A: PtrContainer<_Attribute>
    {
        let p = self.ptr_mut();
        let vb_inherit: VARIANT_BOOL = if inherit {-1} else {0};
        let mut pattrs: *mut SAFEARRAY = ptr::null_mut();
        let hr = match attr {
            Some(attr) => unsafe {
                let t = attr.ptr_mut();
                (*p).GetCustomAttributes(t, vb_inherit, &mut pattrs)
            }, 
            None => unsafe {
                (*p).GetCustomAttributes_2(vb_inherit, &mut pattrs)
            }
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, pattrs, *mut IUnknown, *mut _Attribute, A}, _Assembly )
    }
    
    fn manifest_resource_names(&self) -> Result<Vec<String>>
    {
        let p = self.ptr_mut();
        let mut pnames: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetManifestResourceNames(&mut pnames)
        };
        SUCCEEDED!(hr, SIMPLE_EXTRACT!{BString, pnames, BString}, _Assembly )
    }

    fn files<F>(&self, resource_modules: Option<bool>) -> Result<Vec<F>> 
        where F: PtrContainer<_FileStream>
    {
        let p = self.ptr_mut();
        let mut pfiles: *mut SAFEARRAY = ptr::null_mut();
        let hr = match resource_modules {
            Some(get_modules) => unsafe {
                let vb: VARIANT_BOOL = if get_modules {-1} else {0};
                (*p).GetFiles_2(vb, &mut pfiles)
            }, 
            None => unsafe {
                (*p).GetFiles(&mut pfiles)
            }
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, pfiles, *mut IUnknown, *mut _FileStream, F}, _Assembly)
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
        where V: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt = value.into_variant().into_c_variant();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Equals(vt, &mut vb)
        };
        SUCCEEDED!(hr, vb < 0, _Assembly)
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
    //#[todo("Merge both name methods")]
    fn name<A>(&self) -> Result<A>
        where A: PtrContainer<_AssemblyName>
    {
        let p = self.ptr_mut();
        let mut an: *mut _AssemblyName = ptr::null_mut();
        let hr = unsafe {
            (*p).GetName(&mut an)
        };
        SUCCEEDED!(hr, A::from(an), _Assembly)
    }

    fn name_2<A>(&self, use_code_base_after_shadow_copy: bool) -> Result<A>
        where A: PtrContainer<_AssemblyName>
    {
        let p = self.ptr_mut();
        let mut an: *mut _AssemblyName = ptr::null_mut();
        let hr = unsafe {
            (*p).GetName_2(if use_code_base_after_shadow_copy  {-1} else {0} as VARIANT_BOOL, &mut an)
        };
        SUCCEEDED!(hr, A::from(an), _Assembly)
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
        let mut mi: *mut _MethodInfo = ptr::null_mut();
        let hr = unsafe {
            (*p).get_EntryPoint(&mut mi)
        };
        SUCCEEDED!(hr, M::from(mi),  _Assembly)
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
            (*p).GetType_3(bs.as_sys(), if throw_on_error {-1} else {0} as VARIANT_BOOL, t)
        };
        SUCCEEDED!(hr, T::from(unsafe {*t}),  _Assembly)
    }
    
    fn exported_types<S>(&self) -> Result<Vec<S>> 
        where S: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let mut psa: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetExportedTypes(&mut psa)
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, psa, *mut IUnknown, *mut _Type, S} , _Assembly)
    }

    fn types<S>(&self) -> Result<Vec<S>> 
        where S: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let mut psa: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetTypes(&mut psa)
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, psa, *mut IUnknown, *mut _Type, S}, _Assembly)
    }

    fn manifest_resource_stream<T, S>(&self, t: T, name: String) -> Result<S> 
        where T: PtrContainer<_Type>, 
              S: PtrContainer<_Stream>
    {
        let p = self.ptr_mut();
        let t = t.ptr_mut();
        let bs: BString = From::from(name);
        let mut s: *mut _Stream = ptr::null_mut();
        let hr = unsafe {
            (*p).GetManifestResourceStream(t, bs.as_sys(), &mut s)
        };
        SUCCEEDED!(hr, S::from(s), _Assembly)
    }

    fn manifest_resource_stream_2<S>(&self, name: String) -> Result<S> 
        where S: PtrContainer<_Stream> 
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let mut s: *mut _Stream = ptr::null_mut();
        let hr = unsafe {
            (*p).GetManifestResourceStream_2(bs.as_sys(), &mut s)
        };
        SUCCEEDED!(hr, S::from(s), _Assembly)
    }

    fn file<F>(&self, name: String) -> Result<F> 
        where F: PtrContainer<_FileStream>
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let mut f: *mut _FileStream = ptr::null_mut();
        let hr = unsafe {
            (*p).GetFile(bs.as_sys(), &mut f)
        };
        SUCCEEDED!(hr, F::from(f), _Assembly)
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
        let vb_inherit: VARIANT_BOOL = if inherit {-1}  else {0};
        let mut vb_ret: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).IsDefined(t, vb_inherit, &mut vb_ret)
        };
        SUCCEEDED!(hr, vb_ret < 0, _Assembly )
    }

    fn type_4<T>(&self, name: String, throw_on_error: bool, ignore_case: bool) -> Result<T> 
        where T: PtrContainer<_Type> 
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let vb_throw: VARIANT_BOOL = if throw_on_error {-1} else {0};
        let vb_ignore: VARIANT_BOOL = if ignore_case {-1} else {0};
        let mut t: *mut _Type = ptr::null_mut();
        let hr = unsafe {
            (*p).GetType_4(bs.as_sys(), vb_throw, vb_ignore, &mut t)
        };
        SUCCEEDED!(hr, T::from(t), _Assembly)
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

//#[incomplete]
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
    //#[incomplete]
    fn constructor(&self) {
        unimplemented!()
    }

    fn properties<PI>(&self, binding_attrs: BindingFlags) -> Result<Vec<PI>> 
        where PI: PtrContainer<_PropertyInfo> 
    {
        let p = self.ptr_mut();
        let mut psa: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetProperties(binding_attrs, &mut psa)
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, psa, *mut IUnknown, *mut _PropertyInfo, PI}, _Type)
    }

    fn property<P>(&self, name: String, binding_attrs: BindingFlags) -> Result<P>
        where P: PtrContainer<_PropertyInfo>
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let mut ppi: *mut _PropertyInfo = ptr::null_mut();
        let hr = unsafe {
            (*p).GetProperty(bs.as_sys(), binding_attrs, &mut ppi)
        };
        SUCCEEDED!(hr, P::from(ppi), _Type)
    }

    fn fields<F>(&self, binding_attrs: BindingFlags) -> Result<Vec<F>> 
        where F: PtrContainer<_FieldInfo> 
    {
        let p = self.ptr_mut();
        let mut psa: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetFields(binding_attrs, &mut psa)
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, psa, *mut IUnknown, *mut _FieldInfo, F}, _Type)
    }

    fn field<F>(&self, binding_attrs: BindingFlags) -> Result<Vec<F>> 
        where F: PtrContainer<_FieldInfo> 
    {
        let p = self.ptr_mut();
        let mut psa: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetField(binding_attrs, &mut psa)
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, psa, *mut IUnknown, *mut _FieldInfo, F}, _Type)
    }

    fn methods<M>(&self, binding_attrs: BindingFlags) -> Result<Vec<M>> 
        where M: PtrContainer<_MethodInfo>
    {
        let p = self.ptr_mut();
        let mut psa: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetMethods(binding_attrs, &mut psa)
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, psa, *mut IUnknown, *mut _MethodInfo, M}, _Type)
    }

    //still need to implement GetMethod with binder, types, and modifiers
    //#[todo("GetMethod with binder, types, etc")]
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
        let mut pim: *mut InterfaceMapping = ptr::null_mut();
        let hr = unsafe {
            (*p).GetInterfaceMap(t, &mut pim)
        };
        SUCCEEDED!(hr, WrappedInterfaceMapping::from(unsafe{*pim}), _Type)
    }

    fn instance_of_type(&self, variant: Variant) -> Result<bool> 
    {
        let p = self.ptr_mut();
        let v = variant.into_c_variant();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).IsInstanceOfType(v, &mut vb)
        };
        SUCCEEDED!(hr, vb < 0, _Type)        
    }

    fn assignable_from<T>(&self, test_type: T) -> Result<bool> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let t = test_type.ptr_mut();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).IsAssignableFrom(t, &mut vb)
        };
        SUCCEEDED!(hr, vb < 0, _Type)        
    }

    fn subclass_of<T>(&self, test_type: T) -> Result<bool> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let t = test_type.ptr_mut();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).IsSubclassOf(t, &mut vb)
        };
        SUCCEEDED!(hr, vb < 0, _Type)        
    }

    //#[incomplete]
    fn find_members(&self)
    {
        unimplemented!()
    }

    fn default_members<M>(&self) -> Result<Vec<M>>
        where M: PtrContainer<_MemberInfo>
    {
        let p = self.ptr_mut();
        let mut pm: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetDefaultMembers(&mut pm)
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, pm, *mut IUnknown, *mut _MemberInfo, M}, _Type)
    }

    //#[incomplete]
    fn invoke_member(&self) {
        unimplemented!()
    }

    fn members<M>(&self, binding_attr: BindingFlags) -> Result<Vec<M>> 
        where M: PtrContainer<_MemberInfo>
    {
        let p = self.ptr_mut();
        let mut psa: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetMembers(binding_attr, &mut psa)
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, psa, *mut IUnknown, *mut _MemberInfo, M}, _Type)
    }

    fn member<M>(&self, name: String, member_types: Option<MemberTypes>, binding_flags: BindingFlags) -> Result<Vec<M>>
        where M: PtrContainer<_MemberInfo>
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let mut ppm: *mut SAFEARRAY = ptr::null_mut();
        let hr = match member_types {
            Some(member_types) => unsafe {
                (*p).GetMember(bs.as_sys(), member_types, binding_flags, &mut ppm)
            }, 
            None => unsafe {
                (*p).GetMember_2(bs.as_sys(), binding_flags, &mut ppm)
            }
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, ppm, *mut IUnknown, *mut _MemberInfo, M}, _Type)
    }

    fn nested_type<T>(&self, name: String, binding_flags: BindingFlags) -> Result<T> 
        where T: PtrContainer<_Type> 
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let mut ppt: *mut _Type = ptr::null_mut();
        let hr = unsafe {
            (*p).GetNestedType(bs.as_sys(), binding_flags, &mut ppt)
        };
        SUCCEEDED!(hr, T::from(ppt), _Type)
    }

    fn nested_types<T>(&self, binding_flags: BindingFlags) -> Result<Vec<T>> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let mut psa: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetNestedTypes(binding_flags, &mut psa)
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, psa, *mut IUnknown, *mut _Type, T}, _Type)
    }

    fn events<E>(&self, binding_flags: Option<BindingFlags>) -> Result<Vec<E>> 
        where E: PtrContainer<_EventInfo>
    {
        let p = self.ptr_mut();
        let mut e: *mut SAFEARRAY = ptr::null_mut();

        let hr = match binding_flags {
            Some(flags) => unsafe {
                (*p).GetEvents_2(flags, &mut e)
            }, 
            None => unsafe {
                (*p).GetEvents(&mut e)
            }
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, e, *mut IUnknown, *mut _EventInfo, E}, _Type)
    }

    fn event<E>(&self, name: String, flags: BindingFlags) -> Result<E>
        where E: PtrContainer<_EventInfo> 
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let mut e: *mut _EventInfo = ptr::null_mut();
        let hr = unsafe {
            (*p).GetEvent(bs.as_sys(), flags, &mut e)
        };
        SUCCEEDED!(hr, E::from(e), _Type)
    }

    //#[incomplete]
    fn find_interfaces<TF, T, V, E>(&self, _filter: TF, _criteria: V) -> Result<RSafeArray<T>> 
        where TF: PtrContainer<_TypeFilter>,
              T: PtrContainer<_Type>,
              V: PtrContainer<E>
    {
        unimplemented!();
    }
    fn interfaces<T>(&self) -> Result<Vec<T>> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let mut psa: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetInterfaces(&mut psa)
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, psa, *mut IUnknown, *mut _Type, T}, _Type)
    }

    fn interface<T>(&self, name: String, ignore_case: bool) -> Result<T> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let bs: BString = From::from(name);
        let mut t: *mut _Type = ptr::null_mut();
        let vb: VARIANT_BOOL = if ignore_case{-1} else{0};
        let hr = unsafe {
            (*p).GetInterface(bs.as_sys(), vb, &mut t)
        };
        SUCCEEDED!(hr, T::from(t), _Type)
    }

    fn constructors<C>(&self, binding_attrs: BindingFlags) -> Result<Vec<C>> 
        where C: PtrContainer<_ConstructorInfo>
    {
        let p = self.ptr_mut();
        let mut psa: *mut SAFEARRAY = ptr::null_mut();
        let hr = unsafe {
            (*p).GetConstructors(binding_attrs, &mut psa)
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, psa, *mut IUnknown, *mut _ConstructorInfo, C}, _Type)
    }

    fn defined<T>(&self, attr_type: T, inherit: bool) -> Result<bool>
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let vb: VARIANT_BOOL = if inherit{-1} else {0};
        let attr = attr_type.ptr_mut();
        let mut ret: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).IsDefined(attr, vb, &mut ret)
        };
        SUCCEEDED!(hr, ret < 0, _Type)
    }

    fn custom_attributes<T, A>(&self, inherit: bool, attr_type: Option<T>) -> Result<Vec<A>>
        where A: PtrContainer<_Attribute>, 
              T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let vb: VARIANT_BOOL = if inherit {-1} else {0};
        let mut psa: *mut SAFEARRAY = ptr::null_mut();
        let hr = match attr_type {
            Some(attr) => unsafe {
                let t = attr.ptr_mut();
                (*p).GetCustomAttributes(t, vb, &mut psa)
            },
            None => unsafe {
                (*p).GetCustomAttributes_2(vb, &mut psa)
            }
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, psa, *mut IUnknown, *mut _Attribute, A}, _Type)
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
        where V: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt = value.into_variant().into_c_variant();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Equals(vt, &mut vb)
        };
        SUCCEEDED!(hr, vb < 0, _Type)
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
        SUCCEEDED!(hr, vb < 0, _Type)
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

    //#[incomplete]
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
//#[incomplete]
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
        where V: PtrContainer<TOut>
    {
        let p = self.ptr_mut();
        let vt = value.into_variant().into_c_variant();
        let mut vb: VARIANT_BOOL = 0;
        let hr = unsafe {
            (*p).Equals(vt, &mut vb)
        };
        SUCCEEDED!(hr, vb < 0, _MemberInfo)
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
    
    fn custom_attributes<T, A>(&self, inherit: bool, attr_type: Option<T>) -> Result<Vec<A>>
        where T: PtrContainer<_Type>, 
              A: PtrContainer<_Attribute>
    {
        let p = self.ptr_mut();
        let mut psa: *mut SAFEARRAY = ptr::null_mut();
        let vb: VARIANT_BOOL = if inherit {-1} else {0};
        let hr = match attr_type {
            Some(attr_type) => unsafe {
                let t = attr_type.ptr_mut();
                (*p).GetCustomAttributes(t, vb, &mut psa)
            }, 
            None => unsafe {
                (*p).GetCustomAttributes_2(vb, &mut psa)
            }
        };
        SUCCEEDED!(hr, EXTRACT_VECTOR_FROM_SAFEARRAY!{Unknown, psa, *mut IUnknown, *mut _Attribute, A}, _MemberInfo)
    }

    fn is_defined<T>(&self, attr_type: T, inherit: bool) -> Result<bool> 
        where T: PtrContainer<_Type>
    {
        let p = self.ptr_mut();
        let vb: VARIANT_BOOL = if inherit {-1} else {0};
        let mut ret: VARIANT_BOOL = 0;
        let t = attr_type.ptr_mut();
        let hr = unsafe {
            (*p).IsDefined(t, vb, &mut ret)
        };
        SUCCEEDED!(hr, ret < 0, _MemberInfo)
    }
}
//#[incomplete]
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
//#[incomplete]
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
//#[incomplete]
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
//#[incomplete]
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
//#[incomplete]
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
//#[incomplete]
pub trait EventInfo where Self: PtrContainer<_EventInfo> {
    PROPERTY!{get_ToString _EventInfo { get {to_str(u16)}}}
    PROPERTY!{GetType _EventInfo { get {type_of(_Type)}}}
    PROPERTY!{get_name _EventInfo { get {name(u16)}}}

    PROPERTY!{get_DeclaringType _EventInfo { get {declaring_type(_Type)}}}
    PROPERTY!{get_ReflectedType _EventInfo { get {reflected_type(_Type)}}}
    PROPERTY!{get_IsMulticast _EventInfo { get {multicast(VARIANT_BOOL)}}}
    PROPERTY!{get_EventHandlerType _EventInfo { get {event_handler_type(_Type)}}}
}
//#[incomplete]
pub trait ParameterInfo where Self: PtrContainer<_ParameterInfo> {

}
//#[incomplete]
pub trait Module where Self: PtrContainer<_Module> {
    
}
//#[incomplete]
pub trait AssemblyName where Self: PtrContainer<_AssemblyName> {

}
//#[incomplete]
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
     
    fn into_variant(&self) -> Variant {
        let p = self.ptr_mut();
        let p: *mut IUnknown = unsafe {mem::transmute::<*mut _Type, *mut IUnknown>(p)};
        Variant::from(p)
    }
}

impl Display for ClrType {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        let s = self.to_str().unwrap_or(From::from("ClrType"));
        write!(f, "{:?}", s)
    }
}