// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentVisitor.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/SubDocument.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeEnd.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeStart.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Markup/SmartTag.h>
#include <Aspose.Words.Cpp/Model/Math/OfficeMath.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeStart.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <system/shared_ptr.h>
#include <system/string.h>
#include <system/text/string_builder.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Math;
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
