#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <system/array.h>
#include <system/io/stream.h>

#include "ApiExampleBase.h"
#include "TestData/TestClasses/DocumentTestClass.h"

using namespace ApiExamples::TestData::TestClasses;
using namespace Aspose::Words;

namespace ApiExamples { namespace TestData { namespace TestBuilders {

class DocumentTestBuilder : public ApiExampleBase
{
public:
    DocumentTestBuilder();

    System::SharedPtr<DocumentTestBuilder> WithDocument(System::SharedPtr<Document> doc);
    System::SharedPtr<DocumentTestBuilder> WithDocumentStream(System::SharedPtr<System::IO::Stream> stream);
    System::SharedPtr<DocumentTestBuilder> WithDocumentBytes(System::ArrayPtr<uint8_t> docBytes);
    System::SharedPtr<DocumentTestBuilder> WithDocumentString(System::String docString);
    System::SharedPtr<DocumentTestClass> Build();

protected:
    virtual ~DocumentTestBuilder();

private:
    System::SharedPtr<Document> mDocument;
    System::SharedPtr<System::IO::Stream> mDocumentStream;
    System::ArrayPtr<uint8_t> mDocumentBytes;
    System::String mDocumentString;
};

}}} // namespace ApiExamples::TestData::TestBuilders
