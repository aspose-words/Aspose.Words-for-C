// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExInline.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionCollection.h>
#include <Aspose.Words.Cpp/Model/Revisions/Revision.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
namespace ApiExamples {

namespace gtest_test
{

class ExInline : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExInline> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExInline>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExInline> ExInline::s_instance;

} // namespace gtest_test

void ExInline::InlineRevisions()
{
    //ExStart
    //ExFor:Inline
    //ExFor:Inline.IsDeleteRevision
    //ExFor:Inline.IsFormatRevision
    //ExFor:Inline.IsInsertRevision
    //ExFor:Inline.IsMoveFromRevision
    //ExFor:Inline.IsMoveToRevision
    //ExFor:Inline.ParentParagraph
    //ExFor:Paragraph.Runs
    //ExFor:Revision.ParentNode
    //ExFor:RunCollection
    //ExFor:RunCollection.Item(Int32)
    //ExFor:RunCollection.ToArray
    //ExSummary:Shows how to view revision-related properties of Inline nodes.
    auto doc = MakeObject<Document>(MyDir + u"Revision runs.docx");

    // This document has 6 revisions
    ASSERT_EQ(6, doc->get_Revisions()->get_Count());

    // The parent node of a revision is the run that the revision concerns, which is an Inline node
    auto run = System::DynamicCast<Aspose::Words::Run>(doc->get_Revisions()->idx_get(0)->get_ParentNode());

    // Get the parent paragraph
    SharedPtr<Paragraph> firstParagraph = run->get_ParentParagraph();
    SharedPtr<RunCollection> runs = firstParagraph->get_Runs();

    ASSERT_EQ(6, runs->ToArray()->get_Length());

    // The text in the run at index #2 was typed after revisions were tracked, so it will count as an insert revision
    // The font was changed, so it will also be a format revision
    ASSERT_TRUE(runs->idx_get(2)->get_IsInsertRevision());
    ASSERT_TRUE(runs->idx_get(2)->get_IsFormatRevision());

    // If one node was moved from one place to another while changes were tracked,
    // the node will be placed at the departure location as a "move to revision",
    // and a "move from revision" node will be left behind at the origin, in case we want to reject changes
    // Highlighting text and dragging it to another place with the mouse and cut-and-pasting (but not copy-pasting) both count as "move revisions"
    // The node with the "IsMoveToRevision" flag is the arrival of the move operation, and the node with the "IsMoveFromRevision" flag is the departure point
    ASSERT_TRUE(runs->idx_get(1)->get_IsMoveToRevision());
    ASSERT_TRUE(runs->idx_get(4)->get_IsMoveFromRevision());

    // If an Inline node gets deleted while changes are being tracked, it will leave behind a node with the IsDeleteRevision flag set to true until changes are accepted
    ASSERT_TRUE(runs->idx_get(5)->get_IsDeleteRevision());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExInline, InlineRevisions)
{
    s_instance->InlineRevisions();
}

} // namespace gtest_test

} // namespace ApiExamples
