//use winapi::um::oaidl::SAFEARRAY;

use mscorlib_sys::system::reflection::{_MethodInfo, _Type};
use mscorlib_sys::system::reflection::InterfaceMapping as comInterfaceMapping;

use wrappers::PtrContainer;
use safearray::{SafeArray};
use variant::{WrappedDispatch, PhantomDispatch };

pub struct InterfaceMapping<PtrTarget, PtrInterface, M> 
    where PtrTarget: PtrContainer<_Type>, 
          PtrInterface: PtrContainer<_Type>,
          M: PtrContainer<_MethodInfo>
{
    pub target: PtrTarget, 
    pub interface: PtrInterface,
    pub target_methods: SafeArray<WrappedDispatch, M, PhantomDispatch, _MethodInfo, String>,
    pub interface_methods: SafeArray<WrappedDispatch, M, PhantomDispatch, _MethodInfo, String>,
}

impl<PtrTarget, PtrInterface, M>  From<comInterfaceMapping> for InterfaceMapping<PtrTarget, PtrInterface, M>  
    where PtrTarget: PtrContainer<_Type>, 
          PtrInterface: PtrContainer<_Type>,
          M: PtrContainer<_MethodInfo>
{
    fn from(cim: comInterfaceMapping) -> InterfaceMapping<PtrTarget, PtrInterface, M> {
        InterfaceMapping {
            target: PtrTarget::from(cim.TargetType), 
            interface: PtrInterface::from(cim.interfaceType), 
            target_methods: SafeArray::from(cim.TargetMethods), 
            interface_methods: SafeArray::from(cim.InterfaceMethods)
        }
    }
}