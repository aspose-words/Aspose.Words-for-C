// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExComHelper.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/details/dispose_guard.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/ComHelper.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
namespace ApiExamples {

namespace gtest_test
{

class ExComHelper : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExComHelper> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExComHelper>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExComHelper> ExComHelper::s_instance;

} // namespace gtest_test

void ExComHelper::ComHelper()
{
    //ExStart
    //ExFor:ComHelper
    //ExFor:ComHelper.#ctor
    //ExFor:ComHelper.Open(Stream)
    //ExFor:ComHelper.Open(String)
    //ExSummary:Shows how to open documents using the ComHelper class.
    // If you need to open a document within a COM application,
    // you will need to do so using the ComHelper class as instead of the Document constructor
    auto comHelper = MakeObject<Aspose::Words::ComHelper>();

    // There are two ways of using a ComHelper to open a document
    // 1: Using a filename
    SharedPtr<Document> doc = comHelper->Open(MyDir + u"Document.docx");
    ASSERT_EQ(u"Hello World!", doc->GetText().Trim());

    // 2: Using a Stream
    {
        auto stream = MakeObject<System::IO::FileStream>(MyDir + u"Document.docx", System::IO::FileMode::Open);
        doc = comHelper->Open(stream);
        ASSERT_EQ(u"Hello World!", doc->GetText().Trim());
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExComHelper, ComHelper)
{
    s_instance->ComHelper();
}

} // namespace gtest_test

} // namespace ApiExamples
