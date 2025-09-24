// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#include "HelperClasses/TestBuilders/ImageTestBuilder.h"


using namespace Aspose::Words::ApiExamples::TestData::TestClasses;
namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestBuilders {

RTTI_INFO_IMPL_HASH(3827196177u, ::Aspose::Words::ApiExamples::TestData::TestBuilders::ImageTestBuilder, ThisTypeBaseTypesInfo);

ImageTestBuilder::ImageTestBuilder()
{
    mImageStream = System::IO::Stream::Null;
    mImageBytes = System::MakeArray<uint8_t>(0, 0);
    mImageString = System::String::Empty;
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ImageTestBuilder> ImageTestBuilder::WithImage(System::String imagePath)
{
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ImageTestBuilder> ImageTestBuilder::WithImageStream(System::SharedPtr<System::IO::Stream> imageStream)
{
    mImageStream = imageStream;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ImageTestBuilder> ImageTestBuilder::WithImageBytes(System::ArrayPtr<uint8_t> imageBytes)
{
    mImageBytes = imageBytes;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ImageTestBuilder> ImageTestBuilder::WithImageString(System::String imageString)
{
    mImageString = imageString;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ImageTestClass> ImageTestBuilder::Build()
{
    return System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ImageTestClass>(mImage, mImageStream, mImageBytes, mImageString);
}

ImageTestBuilder::~ImageTestBuilder()
{
}

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
