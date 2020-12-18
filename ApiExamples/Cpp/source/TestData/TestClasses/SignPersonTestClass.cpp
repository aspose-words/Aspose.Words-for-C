#include "TestData/TestClasses/SignPersonTestClass.h"

#include <system/string.h>
#include <system/scope_guard.h>
#include <system/guid.h>
#include <system/array.h>
#include <cstdint>

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

System::Guid SignPersonTestClass::get_PersonId()
{
    return pr_PersonId;
}

void SignPersonTestClass::set_PersonId(System::Guid value)
{
    pr_PersonId = value;
}

System::String SignPersonTestClass::get_Name()
{
    return pr_Name;
}

void SignPersonTestClass::set_Name(System::String value)
{
    pr_Name = value;
}

System::String SignPersonTestClass::get_Position()
{
    return pr_Position;
}

void SignPersonTestClass::set_Position(System::String value)
{
    pr_Position = value;
}

System::ArrayPtr<uint8_t> SignPersonTestClass::get_Image()
{
    return pr_Image;
}

void SignPersonTestClass::set_Image(System::ArrayPtr<uint8_t> value)
{
    pr_Image = value;
}

SignPersonTestClass::SignPersonTestClass(System::Guid guid, System::String name, System::String position, System::ArrayPtr<uint8_t> image)
{
    //Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference
    
    set_PersonId(guid);
    set_Name(name);
    set_Position(position);
    set_Image(image);
}

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
