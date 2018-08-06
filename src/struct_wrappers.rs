//use winapi::um::oaidl::SAFEARRAY;

use mscorlib_sys::system::reflection::{ _Type};
use mscorlib_sys::system::reflection::InterfaceMapping as comInterfaceMapping;

use wrappers::PtrContainer;
use new_safearray::RSafeArray;

pub struct InterfaceMapping<PtrTarget, PtrInterface> 
    where PtrTarget: PtrContainer<_Type>, 
          PtrInterface: PtrContainer<_Type>
{
    pub target: PtrTarget, 
    pub interface: PtrInterface,
    pub target_methods: RSafeArray,
    pub interface_methods: RSafeArray,
}

impl<PtrTarget, PtrInterface>  From<comInterfaceMapping> for InterfaceMapping<PtrTarget, PtrInterface>  
    where PtrTarget: PtrContainer<_Type>, 
          PtrInterface: PtrContainer<_Type>
{
    fn from(cim: comInterfaceMapping) -> InterfaceMapping<PtrTarget, PtrInterface> {
        InterfaceMapping {
            target: PtrTarget::from(cim.TargetType), 
            interface: PtrInterface::from(cim.interfaceType), 
            target_methods: RSafeArray::from(cim.TargetMethods), 
            interface_methods: RSafeArray::from(cim.InterfaceMethods)
        }
    }
}