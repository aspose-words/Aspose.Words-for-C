#include "TestData/TestClasses/MessageTestClass.h"

#include <system/string.h>
#include <system/scope_guard.h>

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

RTTI_INFO_IMPL_HASH(2167543284u, ::ApiExamples::TestData::TestClasses::MessageTestClass, ThisTypeBaseTypesInfo);

System::String MessageTestClass::get_Name()
{
    return pr_Name;
}

void MessageTestClass::set_Name(System::String value)
{
    pr_Name = value;
}

System::String MessageTestClass::get_Message()
{
    return pr_Message;
}

void MessageTestClass::set_Message(System::String value)
{
    pr_Message = value;
}

MessageTestClass::MessageTestClass(System::String name, System::String message)
{
    //Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference

    set_Name(name);
    set_Message(message);
}

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
