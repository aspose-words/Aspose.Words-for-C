#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.

#include <system/io/stream.h>
#include <system/array.h>
#include <drawing/image.h>
#include <cstdint>

namespace Aspose {

namespace Words {

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

    System::SharedPtr<System::Drawing::Image> get_Image() const;
    void set_Image(System::SharedPtr<System::Drawing::Image> value);
    System::SharedPtr<System::IO::Stream> get_ImageStream() const;
    void set_ImageStream(System::SharedPtr<System::IO::Stream> value);
    System::ArrayPtr<uint8_t> get_ImageBytes() const;
    void set_ImageBytes(System::ArrayPtr<uint8_t> value);
    System::String get_ImageString() const;
    void set_ImageString(System::String value);
    
    ImageTestClass(System::SharedPtr<System::Drawing::Image> image, System::SharedPtr<System::IO::Stream> imageStream, System::ArrayPtr<uint8_t> imageBytes, System::String imageString);
    
private:

    System::SharedPtr<System::Drawing::Image> pr_Image;
    System::SharedPtr<System::IO::Stream> pr_ImageStream;
    System::ArrayPtr<uint8_t> pr_ImageBytes;
    System::String pr_ImageString;
    
};

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


