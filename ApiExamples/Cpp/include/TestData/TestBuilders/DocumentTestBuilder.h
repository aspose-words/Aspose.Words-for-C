#pragma once

#include <system/io/stream.h>
#include <system/array.h>
#include <cstdint>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class Document; } }
namespace ApiExamples { namespace TestData { namespace TestClasses { class DocumentTestClass; } } }

namespace ApiExamples {

namespace TestData {

namespace TestBuilders {

class DocumentTestBuilder : public ApiExampleBase
{
    typedef DocumentTestBuilder ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    DocumentTestBuilder();
    
    System::SharedPtr<DocumentTestBuilder> WithDocument(System::SharedPtr<Aspose::Words::Document> doc);
    System::SharedPtr<DocumentTestBuilder> WithDocumentStream(System::SharedPtr<System::IO::Stream> stream);
    System::SharedPtr<DocumentTestBuilder> WithDocumentBytes(System::ArrayPtr<uint8_t> docBytes);
    System::SharedPtr<DocumentTestBuilder> WithDocumentString(System::String docString);
    System::SharedPtr<ApiExamples::TestData::TestClasses::DocumentTestClass> Build();
    
protected:

    virtual ~DocumentTestBuilder();
    
    System::Object::shared_members_type GetSharedMembers() override;
    
private:

    System::SharedPtr<Aspose::Words::Document> mDocument;
    System::SharedPtr<System::IO::Stream> mDocumentStream;
    System::ArrayPtr<uint8_t> mDocumentBytes;
    System::String mDocumentString;
    
};

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples


