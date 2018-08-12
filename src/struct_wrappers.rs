//use winapi::um::oaidl::SAFEARRAY;
use mscorlib_sys::system::reflection::{ _Type};
use mscorlib_sys::system::reflection::InterfaceMapping as comInterfaceMapping;
use mscorlib_sys::system::reflection::_MethodInfo;

use wrappers::PtrContainer;
use new_safearray::RSafeArray;

pub struct InterfaceMapping<PtrTarget, PtrInterface, M> 
    where PtrTarget: PtrContainer<_Type>, 
          PtrInterface: PtrContainer<_Type>, 
          M: PtrContainer<_MethodInfo>
{
    pub target: PtrTarget, 
    pub interface: PtrInterface,
    pub target_methods: RSafeArray<M>,
    pub interface_methods: RSafeArray<M>,
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
            target_methods: RSafeArray::from(cim.TargetMethods), 
            interface_methods: RSafeArray::from(cim.InterfaceMethods)
        }
    }
}

