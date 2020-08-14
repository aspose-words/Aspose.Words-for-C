#pragma once

#include <system/string.h>
#include <system/object.h>

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

class ClientTestClass : public System::Object
{
    typedef ClientTestClass ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
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

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples


