#pragma once

#include <cstdint>
#include <drawing/image.h>
#include <system/array.h>
#include <system/io/stream.h>

#include "ApiExampleBase.h"
#include "TestData/TestClasses/ImageTestClass.h"

using namespace ApiExamples::TestData::TestClasses;

namespace ApiExamples { namespace TestData { namespace TestBuilders {

class ImageTestBuilder : public ApiExampleBase
{
public:
    ImageTestBuilder();

    System::SharedPtr<ImageTestBuilder> WithImage(System::SharedPtr<System::Drawing::Image> image);
    System::SharedPtr<ImageTestBuilder> WithImageStream(System::SharedPtr<System::IO::Stream> imageStream);
    System::SharedPtr<ImageTestBuilder> WithImageBytes(System::ArrayPtr<uint8_t> imageBytes);
    System::SharedPtr<ImageTestBuilder> WithImageString(System::String imageString);
    System::SharedPtr<ImageTestClass> Build();

protected:
    virtual ~ImageTestBuilder();

private:
    System::SharedPtr<System::Drawing::Image> mImage;
    System::SharedPtr<System::IO::Stream> mImageStream;
    System::ArrayPtr<uint8_t> mImageBytes;
    System::String mImageString;
};

}}} // namespace ApiExamples::TestData::TestBuilders
