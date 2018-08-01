use mscorlib_sys::system::reflection::{_Binder, ParameterModifier, _Type};

use safearray::SafeArray;
use wrappers::PtrContainer;
use variant::{WrappedDispatch, PhantomDispatch};

#[allow(dead_code)]
pub struct BindingParameters<T> 
    where T: PtrContainer<_Type>
{
    binder: Box<PtrContainer<_Binder>>,
    types: SafeArray<WrappedDispatch, T, PhantomDispatch, _Type, String>, 
    modifiers: Vec<ParameterModifier>
}