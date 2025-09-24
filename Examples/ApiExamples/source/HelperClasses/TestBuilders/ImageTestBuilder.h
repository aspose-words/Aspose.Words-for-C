#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.

#include <system/io/stream.h>
#include <system/array.h>
#include <drawing/image.h>
#include <cstdint>

#include "HelperClasses/TestClasses/ImageTestClass.h"
#include "ApiExampleBase.h"


using namespace Aspose::Words::ApiExamples::TestData::TestClasses;

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestBuilders {

class ImageTestBuilder : public ApiExampleBase
{
    typedef ImageTestBuilder ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    ImageTestBuilder();
    
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ImageTestBuilder> WithImage(System::String imagePath);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ImageTestBuilder> WithImageStream(System::SharedPtr<System::IO::Stream> imageStream);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ImageTestBuilder> WithImageBytes(System::ArrayPtr<uint8_t> imageBytes);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ImageTestBuilder> WithImageString(System::String imageString);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ImageTestClass> Build();
    
protected:

    virtual ~ImageTestBuilder();
    
private:

    System::SharedPtr<System::Drawing::Image> mImage;
    System::SharedPtr<System::IO::Stream> mImageStream;
    System::ArrayPtr<uint8_t> mImageBytes;
    System::String mImageString;
    
};

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


