#pragma once

#include <system/string.h>
#include <system/object.h>
#include <system/guid.h>
#include <system/array.h>
#include <cstdint>

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

class SignPersonTestClass : public System::Object
{
    typedef SignPersonTestClass ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    System::Guid get_PersonId();
    void set_PersonId(System::Guid value);
    System::String get_Name();
    void set_Name(System::String value);
    System::String get_Position();
    void set_Position(System::String value);
    System::ArrayPtr<uint8_t> get_Image();
    void set_Image(System::ArrayPtr<uint8_t> value);
    
    SignPersonTestClass(System::Guid guid, System::String name, System::String position, System::ArrayPtr<uint8_t> image);
    
protected:

    System::Object::shared_members_type GetSharedMembers() override;
    
private:

    System::Guid pr_PersonId;
    System::String pr_Name;
    System::String pr_Position;
    System::ArrayPtr<uint8_t> pr_Image;
    
};

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples


