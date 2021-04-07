#include "TestData/TestBuilders/ImageTestBuilder.h"

using namespace ApiExamples::TestData::TestClasses;
namespace ApiExamples { namespace TestData { namespace TestBuilders {

ImageTestBuilder::ImageTestBuilder()
{
    mImage = System::Drawing::Image::FromFile(ImageDir + u"Transparent background logo.png");
    mImageStream = System::IO::Stream::Null;
    mImageBytes = System::MakeArray<uint8_t>(0, 0);
    mImageString = System::String::Empty;
}

System::SharedPtr<ImageTestBuilder> ImageTestBuilder::WithImage(System::SharedPtr<System::Drawing::Image> image)
{
    mImage = image;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<ImageTestBuilder> ImageTestBuilder::WithImageStream(System::SharedPtr<System::IO::Stream> imageStream)
{
    mImageStream = imageStream;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<ImageTestBuilder> ImageTestBuilder::WithImageBytes(System::ArrayPtr<uint8_t> imageBytes)
{
    mImageBytes = imageBytes;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<ImageTestBuilder> ImageTestBuilder::WithImageString(System::String imageString)
{
    mImageString = imageString;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<ImageTestClass> ImageTestBuilder::Build()
{
    return System::MakeObject<ImageTestClass>(mImage, mImageStream, mImageBytes, mImageString);
}

ImageTestBuilder::~ImageTestBuilder()
{
}

}}} // namespace ApiExamples::TestData::TestBuilders
