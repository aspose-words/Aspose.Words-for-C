// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#include "HelperClasses/TestBuilders/DocumentTestBuilder.h"


using namespace Aspose::Words::ApiExamples::TestData::TestClasses;
namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestBuilders {

RTTI_INFO_IMPL_HASH(3030401803u, ::Aspose::Words::ApiExamples::TestData::TestBuilders::DocumentTestBuilder, ThisTypeBaseTypesInfo);

DocumentTestBuilder::DocumentTestBuilder()
{
    mDocument = System::MakeObject<Aspose::Words::Document>();
    mDocumentStream = System::IO::Stream::Null;
    mDocumentBytes = System::MakeArray<uint8_t>(0, 0);
    mDocumentString = System::String::Empty;
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::DocumentTestBuilder> DocumentTestBuilder::WithDocument(System::SharedPtr<Aspose::Words::Document> doc)
{
    mDocument = doc;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::DocumentTestBuilder> DocumentTestBuilder::WithDocumentStream(System::SharedPtr<System::IO::Stream> stream)
{
    mDocumentStream = stream;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::DocumentTestBuilder> DocumentTestBuilder::WithDocumentBytes(System::ArrayPtr<uint8_t> docBytes)
{
    mDocumentBytes = docBytes;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::DocumentTestBuilder> DocumentTestBuilder::WithDocumentString(System::String docString)
{
    mDocumentString = docString;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::DocumentTestClass> DocumentTestBuilder::Build()
{
    return System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::DocumentTestClass>(mDocument, mDocumentStream, mDocumentBytes, mDocumentString);
}

DocumentTestBuilder::~DocumentTestBuilder()
{
}

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
