// struct_wrappers.rs - MIT License
//  Copyright (c) 2018 Tyler Laing (ZerothLaw)
// 
//  Permission is hereby granted, free of charge, to any person obtaining a copy
//  of this software and associated documentation files (the "Software"), to deal
//  in the Software without restriction, including without limitation the rights
//  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//  copies of the Software, and to permit persons to whom the Software is
//  furnished to do so, subject to the following conditions:
// 
//  The above copyright notice and this permission notice shall be included in all
//  copies or substantial portions of the Software.
// 
//  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//  SOFTWARE.

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

