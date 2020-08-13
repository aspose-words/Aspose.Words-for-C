#pragma once

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/io/stream.h>
#include <system/array.h>
#include <drawing/image.h>
#include <cstdint>

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

class ImageTestClass : public System::Object
{
    typedef ImageTestClass ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    System::SharedPtr<System::Drawing::Image> get_Image();
    void set_Image(System::SharedPtr<System::Drawing::Image> value);
    System::SharedPtr<System::IO::Stream> get_ImageStream();
    void set_ImageStream(System::SharedPtr<System::IO::Stream> value);
    System::ArrayPtr<uint8_t> get_ImageBytes();
    void set_ImageBytes(System::ArrayPtr<uint8_t> value);
    System::String get_ImageString();
    void set_ImageString(System::String value);
    
    ImageTestClass(System::SharedPtr<System::Drawing::Image> image, System::SharedPtr<System::IO::Stream> imageStream, System::ArrayPtr<uint8_t> imageBytes, System::String imageString);
    
protected:

    System::Object::shared_members_type GetSharedMembers() override;
    
private:

    System::SharedPtr<System::Drawing::Image> pr_Image;
    System::SharedPtr<System::IO::Stream> pr_ImageStream;
    System::ArrayPtr<uint8_t> pr_ImageBytes;
    System::String pr_ImageString;
    
};

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples


