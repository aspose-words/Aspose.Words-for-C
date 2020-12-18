#include "TestData/TestClasses/ImageTestClass.h"

#include <system/string.h>
#include <system/scope_guard.h>
#include <system/io/stream.h>
#include <system/array.h>
#include <drawing/image.h>
#include <cstdint>

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

System::SharedPtr<System::Drawing::Image> ImageTestClass::get_Image()
{
    return pr_Image;
}

void ImageTestClass::set_Image(System::SharedPtr<System::Drawing::Image> value)
{
    pr_Image = value;
}

System::SharedPtr<System::IO::Stream> ImageTestClass::get_ImageStream()
{
    return pr_ImageStream;
}

void ImageTestClass::set_ImageStream(System::SharedPtr<System::IO::Stream> value)
{
    pr_ImageStream = value;
}

System::ArrayPtr<uint8_t> ImageTestClass::get_ImageBytes()
{
    return pr_ImageBytes;
}

void ImageTestClass::set_ImageBytes(System::ArrayPtr<uint8_t> value)
{
    pr_ImageBytes = value;
}

System::String ImageTestClass::get_ImageString()
{
    return pr_ImageString;
}

void ImageTestClass::set_ImageString(System::String value)
{
    pr_ImageString = value;
}

ImageTestClass::ImageTestClass(System::SharedPtr<System::Drawing::Image> image, System::SharedPtr<System::IO::Stream> imageStream, System::ArrayPtr<uint8_t> imageBytes, System::String imageString)
{
    //Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference
    
    set_Image(image);
    set_ImageStream(imageStream);
    set_ImageBytes(imageBytes);
    set_ImageString(imageString);
}

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
