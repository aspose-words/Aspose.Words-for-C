#pragma once

#include <system/string.h>
#include <system/object.h>

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

class MessageTestClass : public System::Object
{
    typedef MessageTestClass ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
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

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples


