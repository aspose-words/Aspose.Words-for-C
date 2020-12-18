// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExParagraph.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <system/string.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
namespace ApiExamples { namespace gtest_test {

class ExParagraph : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExParagraph> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExParagraph>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExParagraph> ExParagraph::s_instance;

TEST_F(ExParagraph, DocumentBuilderInsertParagraph)
{
    s_instance->DocumentBuilderInsertParagraph();
}

TEST_F(ExParagraph, InsertField)
{
    s_instance->InsertField();
}

TEST_F(ExParagraph, InsertFieldBeforeTextInParagraph)
{
    s_instance->InsertFieldBeforeTextInParagraph();
}

TEST_F(ExParagraph, InsertFieldAfterTextInParagraph)
{
    s_instance->InsertFieldAfterTextInParagraph();
}

TEST_F(ExParagraph, InsertFieldBeforeTextInParagraphWithoutUpdateField)
{
    s_instance->InsertFieldBeforeTextInParagraphWithoutUpdateField();
}

TEST_F(ExParagraph, InsertFieldAfterTextInParagraphWithoutUpdateField)
{
    s_instance->InsertFieldAfterTextInParagraphWithoutUpdateField();
}

TEST_F(ExParagraph, InsertFieldWithoutSeparator)
{
    s_instance->InsertFieldWithoutSeparator();
}

TEST_F(ExParagraph, InsertFieldBeforeParagraphWithoutDocumentAuthor)
{
    s_instance->InsertFieldBeforeParagraphWithoutDocumentAuthor();
}

TEST_F(ExParagraph, InsertFieldAfterParagraphWithoutChangingDocumentAuthor)
{
    s_instance->InsertFieldAfterParagraphWithoutChangingDocumentAuthor();
}

TEST_F(ExParagraph, InsertFieldBeforeRunText)
{
    s_instance->InsertFieldBeforeRunText();
}

TEST_F(ExParagraph, InsertFieldAfterRunText)
{
    s_instance->InsertFieldAfterRunText();
}

TEST_F(ExParagraph, InsertFieldEmptyParagraphWithoutUpdateField)
{
    s_instance->InsertFieldEmptyParagraphWithoutUpdateField();
}

TEST_F(ExParagraph, InsertFieldEmptyParagraphWithUpdateField)
{
    s_instance->InsertFieldEmptyParagraphWithUpdateField();
}

TEST_F(ExParagraph, CompositeNodeChildren)
{
    s_instance->CompositeNodeChildren();
}

TEST_F(ExParagraph, RevisionHistory)
{
    s_instance->RevisionHistory();
}

TEST_F(ExParagraph, GetFormatRevision)
{
    s_instance->GetFormatRevision();
}

TEST_F(ExParagraph, GetFrameProperties)
{
    s_instance->GetFrameProperties();
}

TEST_F(ExParagraph, IsRevision)
{
    s_instance->IsRevision();
}

TEST_F(ExParagraph, BreakIsStyleSeparator)
{
    s_instance->BreakIsStyleSeparator();
}

TEST_F(ExParagraph, TabStops)
{
    s_instance->TabStops();
}

TEST_F(ExParagraph, JoinRuns)
{
    s_instance->JoinRuns();
}

}} // namespace ApiExamples::gtest_test
