#pragma once

#include <system/io/stream.h>
#include <system/array.h>
#include <drawing/image.h>
#include <cstdint>

#include "ApiExampleBase.h"

namespace ApiExamples { namespace TestData { namespace TestClasses { class ImageTestClass; } } }

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
    
    System::SharedPtr<ImageTestBuilder> WithImage(System::SharedPtr<System::Drawing::Image> image);
    System::SharedPtr<ImageTestBuilder> WithImageStream(System::SharedPtr<System::IO::Stream> imageStream);
    System::SharedPtr<ImageTestBuilder> WithImageBytes(System::ArrayPtr<uint8_t> imageBytes);
    System::SharedPtr<ImageTestBuilder> WithImageString(System::String imageString);
    System::SharedPtr<ApiExamples::TestData::TestClasses::ImageTestClass> Build();
    
protected:

    virtual ~ImageTestBuilder();
    
    System::Object::shared_members_type GetSharedMembers() override;
    
private:

    System::SharedPtr<System::Drawing::Image> mImage;
    System::SharedPtr<System::IO::Stream> mImageStream;
    System::ArrayPtr<uint8_t> mImageBytes;
    System::String mImageString;
    
};

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples


