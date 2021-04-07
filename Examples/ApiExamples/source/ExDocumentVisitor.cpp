// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentVisitor.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Math;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;
namespace ApiExamples { namespace gtest_test {

class ExDocumentVisitor : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExDocumentVisitor> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExDocumentVisitor>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExDocumentVisitor> ExDocumentVisitor::s_instance;

TEST_F(ExDocumentVisitor, DocStructureToText)
{
    s_instance->DocStructureToText();
}

TEST_F(ExDocumentVisitor, TableToText)
{
    s_instance->TableToText();
}

TEST_F(ExDocumentVisitor, CommentsToText)
{
    s_instance->CommentsToText();
}

TEST_F(ExDocumentVisitor, FieldToText)
{
    s_instance->FieldToText();
}

TEST_F(ExDocumentVisitor, HeaderFooterToText)
{
    s_instance->HeaderFooterToText();
}

TEST_F(ExDocumentVisitor, EditableRangeToText)
{
    s_instance->EditableRangeToText();
}

TEST_F(ExDocumentVisitor, FootnoteToText)
{
    s_instance->FootnoteToText();
}

TEST_F(ExDocumentVisitor, OfficeMathToText)
{
    s_instance->OfficeMathToText();
}

TEST_F(ExDocumentVisitor, SmartTagToText)
{
    s_instance->SmartTagToText();
}

TEST_F(ExDocumentVisitor, StructuredDocumentTagToText)
{
    s_instance->StructuredDocumentTagToText();
}

}} // namespace ApiExamples::gtest_test
