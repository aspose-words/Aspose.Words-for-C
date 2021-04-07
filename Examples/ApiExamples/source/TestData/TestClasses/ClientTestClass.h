#pragma once

#include <system/string.h>

namespace ApiExamples { namespace TestData { namespace TestClasses {

class ClientTestClass : public System::Object
{
public:
    System::String get_Name();
    void set_Name(System::String value);
    System::String get_Country();
    void set_Country(System::String value);
    System::String get_LocalAddress();
    void set_LocalAddress(System::String value);

private:
    System::String pr_Name;
    System::String pr_Country;
    System::String pr_LocalAddress;
};

}}} // namespace ApiExamples::TestData::TestClasses
