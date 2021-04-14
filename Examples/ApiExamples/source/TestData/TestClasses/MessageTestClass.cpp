#include "TestData/TestClasses/MessageTestClass.h"

namespace ApiExamples { namespace TestData { namespace TestClasses {

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
    set_Name(name);
    set_Message(message);
}

}}} // namespace ApiExamples::TestData::TestClasses
