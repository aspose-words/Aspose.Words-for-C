#pragma once

#include <system/string.h>

namespace ApiExamples { namespace TestData { namespace TestClasses {

class MessageTestClass : public System::Object
{
public:
    System::String get_Name();
    void set_Name(System::String value);
    System::String get_Message();
    void set_Message(System::String value);

    MessageTestClass(System::String name, System::String message);

private:
    System::String pr_Name;
    System::String pr_Message;
};

}}} // namespace ApiExamples::TestData::TestClasses
