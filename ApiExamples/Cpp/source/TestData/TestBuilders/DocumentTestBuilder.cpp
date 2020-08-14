#include "TestData/TestBuilders/DocumentTestBuilder.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/io/stream.h>
#include <system/array.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "TestData/TestClasses/DocumentTestClass.h"

using namespace ApiExamples::TestData::TestClasses;
using namespace Aspose::Words;
namespace ApiExamples {

namespace TestData {

namespace TestBuilders {

RTTI_INFO_IMPL_HASH(3789440965u, ::ApiExamples::TestData::TestBuilders::DocumentTestBuilder, ThisTypeBaseTypesInfo);

DocumentTestBuilder::DocumentTestBuilder()
{
    mDocument = System::MakeObject<Document>();
    mDocumentStream = System::IO::Stream::Null;
    mDocumentBytes = System::MakeArray<uint8_t>(0, 0);
    mDocumentString = System::String::Empty;
}

System::SharedPtr<DocumentTestBuilder> DocumentTestBuilder::WithDocument(System::SharedPtr<Aspose::Words::Document> doc)
{
    mDocument = doc;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<DocumentTestBuilder> DocumentTestBuilder::WithDocumentStream(System::SharedPtr<System::IO::Stream> stream)
{
    mDocumentStream = stream;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<DocumentTestBuilder> DocumentTestBuilder::WithDocumentBytes(System::ArrayPtr<uint8_t> docBytes)
{
    mDocumentBytes = docBytes;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<DocumentTestBuilder> DocumentTestBuilder::WithDocumentString(System::String docString)
{
    mDocumentString = docString;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<ApiExamples::TestData::TestClasses::DocumentTestClass> DocumentTestBuilder::Build()
{
    return System::MakeObject<DocumentTestClass>(mDocument, mDocumentStream, mDocumentBytes, mDocumentString);
}

DocumentTestBuilder::~DocumentTestBuilder()
{
}

System::Object::shared_members_type ApiExamples::TestData::TestBuilders::DocumentTestBuilder::GetSharedMembers()
{
    auto result = ApiExampleBase::GetSharedMembers();

    result.Add("ApiExamples::TestData::TestBuilders::DocumentTestBuilder::mDocument", this->mDocument);
    result.Add("ApiExamples::TestData::TestBuilders::DocumentTestBuilder::mDocumentStream", this->mDocumentStream);
    result.Add("ApiExamples::TestData::TestBuilders::DocumentTestBuilder::mDocumentBytes", this->mDocumentBytes);

    return result;
}

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples
