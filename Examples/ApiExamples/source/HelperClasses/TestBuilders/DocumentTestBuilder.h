#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.

#include <system/io/stream.h>
#include <system/array.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "HelperClasses/TestClasses/DocumentTestClass.h"
#include "ApiExampleBase.h"


using namespace Aspose::Words::ApiExamples::TestData::TestClasses;

namespace Aspose {

namespace Words {

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
    
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::DocumentTestBuilder> WithDocument(System::SharedPtr<Aspose::Words::Document> doc);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::DocumentTestBuilder> WithDocumentStream(System::SharedPtr<System::IO::Stream> stream);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::DocumentTestBuilder> WithDocumentBytes(System::ArrayPtr<uint8_t> docBytes);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::DocumentTestBuilder> WithDocumentString(System::String docString);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::DocumentTestClass> Build();
    
protected:

    virtual ~DocumentTestBuilder();
    
private:

    System::SharedPtr<Aspose::Words::Document> mDocument;
    System::SharedPtr<System::IO::Stream> mDocumentStream;
    System::ArrayPtr<uint8_t> mDocumentBytes;
    System::String mDocumentString;
    
};

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


