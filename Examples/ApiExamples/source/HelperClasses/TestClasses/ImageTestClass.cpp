// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#include "HelperClasses/TestClasses/ImageTestClass.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

RTTI_INFO_IMPL_HASH(1035042566u, ::Aspose::Words::ApiExamples::TestData::TestClasses::ImageTestClass, ThisTypeBaseTypesInfo);

System::SharedPtr<System::Drawing::Image> ImageTestClass::get_Image() const
{
    return pr_Image;
}

void ImageTestClass::set_Image(System::SharedPtr<System::Drawing::Image> value)
{
    pr_Image = value;
}

System::SharedPtr<System::IO::Stream> ImageTestClass::get_ImageStream() const
{
    return pr_ImageStream;
}

void ImageTestClass::set_ImageStream(System::SharedPtr<System::IO::Stream> value)
{
    pr_ImageStream = value;
}

System::ArrayPtr<uint8_t> ImageTestClass::get_ImageBytes() const
{
    return pr_ImageBytes;
}

void ImageTestClass::set_ImageBytes(System::ArrayPtr<uint8_t> value)
{
    pr_ImageBytes = value;
}

System::String ImageTestClass::get_ImageString() const
{
    return pr_ImageString;
}

void ImageTestClass::set_ImageString(System::String value)
{
    pr_ImageString = value;
}

ImageTestClass::ImageTestClass(System::SharedPtr<System::Drawing::Image> image, System::SharedPtr<System::IO::Stream> imageStream, System::ArrayPtr<uint8_t> imageBytes, System::String imageString)
{
    set_Image(image);
    set_ImageStream(imageStream);
    set_ImageBytes(imageBytes);
    set_ImageString(imageString);
}

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
