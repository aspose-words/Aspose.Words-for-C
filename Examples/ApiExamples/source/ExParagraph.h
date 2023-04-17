#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Fields/FieldStart.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/FrameFormat.h>
#include <Aspose.Words.Cpp/HeightRule.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Properties/DocumentProperty.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Revision.h>
#include <Aspose.Words.Cpp/RevisionCollection.h>
#include <Aspose.Words.Cpp/RevisionType.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/TabAlignment.h>
#include <Aspose.Words.Cpp/TabLeader.h>
#include <Aspose.Words.Cpp/TabStop.h>
#include <Aspose.Words.Cpp/TabStopCollection.h>
#include <Aspose.Words.Cpp/Underline.h>
#include <drawing/color.h>
#include <system/collections/ienumerable.h>
#include <system/date_time.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/timespan.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Drawing;

namespace ApiExamples {

class ExParagraph : public ApiExampleBase
{
public:
    void DocumentBuilderInsertParagraph()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertParagraph
        //ExFor:ParagraphFormat.FirstLineIndent
        //ExFor:ParagraphFormat.Alignment
        //ExFor:ParagraphFormat.KeepTogether
        //ExFor:ParagraphFormat.AddSpaceBetweenFarEastAndAlpha
        //ExFor:ParagraphFormat.AddSpaceBetweenFarEastAndDigit
        //ExFor:Paragraph.IsEndOfDocument
        //ExSummary:Shows how to insert a paragraph into the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Font> font = builder->get_Font();
        font->set_Size(16);
        font->set_Bold(true);
        font->set_Color(System::Drawing::Color::get_Blue());
        font->set_Name(u"Arial");
        font->set_Underline(Underline::Dash);

        SharedPtr<ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
        paragraphFormat->set_FirstLineIndent(8);
        paragraphFormat->set_Alignment(ParagraphAlignment::Justify);
        paragraphFormat->set_AddSpaceBetweenFarEastAndAlpha(true);
        paragraphFormat->set_AddSpaceBetweenFarEastAndDigit(true);
        paragraphFormat->set_KeepTogether(true);

        // The "Writeln" method ends the paragraph after appending text
        // and then starts a new line, adding a new paragraph.
        builder->Writeln(u"Hello world!");

        ASSERT_TRUE(builder->get_CurrentParagraph()->get_IsEndOfDocument());
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        SharedPtr<Paragraph> paragraph = doc->get_FirstSection()->get_Body()->get_FirstParagraph();

        ASPOSE_ASSERT_EQ(8, paragraph->get_ParagraphFormat()->get_FirstLineIndent());
        ASSERT_EQ(ParagraphAlignment::Justify, paragraph->get_ParagraphFormat()->get_Alignment());
        ASSERT_TRUE(paragraph->get_ParagraphFormat()->get_AddSpaceBetweenFarEastAndAlpha());
        ASSERT_TRUE(paragraph->get_ParagraphFormat()->get_AddSpaceBetweenFarEastAndDigit());
        ASSERT_TRUE(paragraph->get_ParagraphFormat()->get_KeepTogether());
        ASSERT_EQ(u"Hello world!", paragraph->GetText().Trim());

        SharedPtr<Font> runFont = paragraph->get_Runs()->idx_get(0)->get_Font();

        ASPOSE_ASSERT_EQ(16.0, runFont->get_Size());
        ASSERT_TRUE(runFont->get_Bold());
        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), runFont->get_Color().ToArgb());
        ASSERT_EQ(u"Arial", runFont->get_Name());
        ASSERT_EQ(Underline::Dash, runFont->get_Underline());
    }

    void AppendField()
    {
        //ExStart
        //ExFor:Paragraph.AppendField(FieldType, Boolean)
        //ExFor:Paragraph.AppendField(String)
        //ExFor:Paragraph.AppendField(String, String)
        //ExSummary:Shows various ways of appending fields to a paragraph.
        auto doc = MakeObject<Document>();
        SharedPtr<Paragraph> paragraph = doc->get_FirstSection()->get_Body()->get_FirstParagraph();

        // Below are three ways of appending a field to the end of a paragraph.
        // 1 -  Append a DATE field using a field type, and then update it:
        paragraph->AppendField(FieldType::FieldDate, true);

        // 2 -  Append a TIME field using a field code:
        paragraph->AppendField(u" TIME  \\@ \"HH:mm:ss\" ");

        // 3 -  Append a QUOTE field using a field code, and get it to display a placeholder value:
        paragraph->AppendField(u" QUOTE \"Real value\"", u"Placeholder value");

        ASSERT_EQ(u"Placeholder value", doc->get_Range()->get_Fields()->idx_get(2)->get_Result());

        // This field will display its placeholder value until we update it.
        doc->UpdateFields();

        ASSERT_EQ(u"Real value", doc->get_Range()->get_Fields()->idx_get(2)->get_Result());

        doc->Save(ArtifactsDir + u"Paragraph.AppendField.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Paragraph.AppendField.docx");

        TestUtil::VerifyField(FieldType::FieldDate, u" DATE ", System::DateTime::get_Now(), doc->get_Range()->get_Fields()->idx_get(0),
                              System::TimeSpan(0, 0, 0, 0));
        TestUtil::VerifyField(FieldType::FieldTime, u" TIME  \\@ \"HH:mm:ss\" ", System::DateTime::get_Now(), doc->get_Range()->get_Fields()->idx_get(1),
                              System::TimeSpan(0, 0, 0, 5));
        TestUtil::VerifyField(FieldType::FieldQuote, u" QUOTE \"Real value\"", u"Real value", doc->get_Range()->get_Fields()->idx_get(2));
    }

    void InsertField()
    {
        //ExStart
        //ExFor:Paragraph.InsertField(string, Node, bool)
        //ExFor:Paragraph.InsertField(FieldType, bool, Node, bool)
        //ExFor:Paragraph.InsertField(string, string, Node, bool)
        //ExSummary:Shows various ways of adding fields to a paragraph.
        auto doc = MakeObject<Document>();
        SharedPtr<Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();

        // Below are three ways of inserting a field into a paragraph.
        // 1 -  Insert an AUTHOR field into a paragraph after one of the paragraph's child nodes:
        auto run = MakeObject<Run>(doc);
        run->set_Text(u"This run was written by ");
        para->AppendChild(run);

        doc->get_BuiltInDocumentProperties()->idx_get(u"Author")->set_Value(System::ObjectExt::Box<String>(u"John Doe"));
        para->InsertField(FieldType::FieldAuthor, true, run, true);

        // 2 -  Insert a QUOTE field after one of the paragraph's child nodes:
        run = MakeObject<Run>(doc);
        run->set_Text(u".");
        para->AppendChild(run);

        SharedPtr<Field> field = para->InsertField(u" QUOTE \" Real value\" ", run, true);

        // 3 -  Insert a QUOTE field before one of the paragraph's child nodes,
        // and get it to display a placeholder value:
        para->InsertField(u" QUOTE \" Real value.\"", u" Placeholder value.", field->get_Start(), false);

        ASSERT_EQ(u" Placeholder value.", doc->get_Range()->get_Fields()->idx_get(1)->get_Result());

        // This field will display its placeholder value until we update it.
        doc->UpdateFields();

        ASSERT_EQ(u" Real value.", doc->get_Range()->get_Fields()->idx_get(1)->get_Result());

        doc->Save(ArtifactsDir + u"Paragraph.InsertField.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Paragraph.InsertField.docx");

        TestUtil::VerifyField(FieldType::FieldAuthor, u" AUTHOR ", u"John Doe", doc->get_Range()->get_Fields()->idx_get(0));
        TestUtil::VerifyField(FieldType::FieldQuote, u" QUOTE \" Real value.\"", u" Real value.", doc->get_Range()->get_Fields()->idx_get(1));
        TestUtil::VerifyField(FieldType::FieldQuote, u" QUOTE \" Real value\" ", u" Real value", doc->get_Range()->get_Fields()->idx_get(2));
    }

    void InsertFieldBeforeTextInParagraph()
    {
        SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

        InsertFieldUsingFieldCode(doc, u" AUTHOR ", nullptr, false, 1);

        ASSERT_EQ(u"\u0013 AUTHOR \u0014Test Author\u0015Hello World!\r", DocumentHelper::GetParagraphText(doc, 1));
    }

    void InsertFieldAfterTextInParagraph()
    {
        String date = System::DateTime::get_Today().ToString(u"d");

        SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

        InsertFieldUsingFieldCode(doc, u" DATE ", nullptr, true, 1);

        ASSERT_EQ(String::Format(u"Hello World!\u0013 DATE \u0014{0}\u0015\r", date), DocumentHelper::GetParagraphText(doc, 1));
    }

    void InsertFieldBeforeTextInParagraphWithoutUpdateField()
    {
        SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

        InsertFieldUsingFieldType(doc, FieldType::FieldAuthor, false, nullptr, false, 1);

        ASSERT_EQ(u"\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper::GetParagraphText(doc, 1));
    }

    void InsertFieldAfterTextInParagraphWithoutUpdateField()
    {
        SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

        InsertFieldUsingFieldType(doc, FieldType::FieldAuthor, false, nullptr, true, 1);

        ASSERT_EQ(u"Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper::GetParagraphText(doc, 1));
    }

    void InsertFieldWithoutSeparator()
    {
        SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

        InsertFieldUsingFieldType(doc, FieldType::FieldListNum, true, nullptr, false, 1);

        ASSERT_EQ(u"\u0013 LISTNUM \u0015Hello World!\r", DocumentHelper::GetParagraphText(doc, 1));
    }

    void InsertFieldBeforeParagraphWithoutDocumentAuthor()
    {
        SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();
        doc->get_BuiltInDocumentProperties()->set_Author(u"");

        InsertFieldUsingFieldCodeFieldString(doc, u" AUTHOR ", nullptr, nullptr, false, 1);

        ASSERT_EQ(u"\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper::GetParagraphText(doc, 1));
    }

    void InsertFieldAfterParagraphWithoutChangingDocumentAuthor()
    {
        SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

        InsertFieldUsingFieldCodeFieldString(doc, u" AUTHOR ", nullptr, nullptr, true, 1);

        ASSERT_EQ(u"Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper::GetParagraphText(doc, 1));
    }

    void InsertFieldBeforeRunText()
    {
        SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

        // Add some text into the paragraph
        SharedPtr<Run> run = DocumentHelper::InsertNewRun(doc, u" Hello World!", 1);

        InsertFieldUsingFieldCodeFieldString(doc, u" AUTHOR ", u"Test Field Value", run, false, 1);

        ASSERT_EQ(u"Hello World!\u0013 AUTHOR \u0014Test Field Value\u0015 Hello World!\r", DocumentHelper::GetParagraphText(doc, 1));
    }

    void InsertFieldAfterRunText()
    {
        SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

        // Add some text into the paragraph
        SharedPtr<Run> run = DocumentHelper::InsertNewRun(doc, u" Hello World!", 1);

        InsertFieldUsingFieldCodeFieldString(doc, u" AUTHOR ", u"", run, true, 1);

        ASSERT_EQ(u"Hello World! Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper::GetParagraphText(doc, 1));
    }

    void InsertFieldEmptyParagraphWithoutUpdateField()
    {
        SharedPtr<Document> doc = DocumentHelper::CreateDocumentWithoutDummyText();

        InsertFieldUsingFieldType(doc, FieldType::FieldAuthor, false, nullptr, false, 1);

        ASSERT_EQ(u"\u0013 AUTHOR \u0014\u0015\f", DocumentHelper::GetParagraphText(doc, 1));
    }

    void InsertFieldEmptyParagraphWithUpdateField()
    {
        SharedPtr<Document> doc = DocumentHelper::CreateDocumentWithoutDummyText();

        InsertFieldUsingFieldType(doc, FieldType::FieldAuthor, true, nullptr, false, 0);

        ASSERT_EQ(u"\u0013 AUTHOR \u0014Test Author\u0015\r", DocumentHelper::GetParagraphText(doc, 0));
    }

    void CompositeNodeChildren()
    {
        //ExStart
        //ExFor:CompositeNode.Count
        //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
        //ExFor:CompositeNode.InsertAfter(Node, Node)
        //ExFor:CompositeNode.InsertBefore(Node, Node)
        //ExFor:CompositeNode.PrependChild(Node)
        //ExFor:Paragraph.GetText
        //ExFor:Run
        //ExSummary:Shows how to add, update and delete child nodes in a CompositeNode's collection of children.
        auto doc = MakeObject<Document>();

        // An empty document, by default, has one paragraph.
        ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->get_Paragraphs()->get_Count());

        // Composite nodes such as our paragraph can contain other composite and inline nodes as children.
        SharedPtr<Paragraph> paragraph = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
        auto paragraphText = MakeObject<Run>(doc, u"Initial text. ");
        paragraph->AppendChild(paragraphText);

        // Create three more run nodes.
        auto run1 = MakeObject<Run>(doc, u"Run 1. ");
        auto run2 = MakeObject<Run>(doc, u"Run 2. ");
        auto run3 = MakeObject<Run>(doc, u"Run 3. ");

        // The document body will not display these runs until we insert them into a composite node
        // that itself is a part of the document's node tree, as we did with the first run.
        // We can determine where the text contents of nodes that we insert
        // appears in the document by specifying an insertion location relative to another node in the paragraph.
        ASSERT_EQ(u"Initial text.", paragraph->GetText().Trim());

        // Insert the second run into the paragraph in front of the initial run.
        paragraph->InsertBefore(run2, paragraphText);

        ASSERT_EQ(u"Run 2. Initial text.", paragraph->GetText().Trim());

        // Insert the third run after the initial run.
        paragraph->InsertAfter(run3, paragraphText);

        ASSERT_EQ(u"Run 2. Initial text. Run 3.", paragraph->GetText().Trim());

        // Insert the first run to the start of the paragraph's child nodes collection.
        paragraph->PrependChild(run1);

        ASSERT_EQ(u"Run 1. Run 2. Initial text. Run 3.", paragraph->GetText().Trim());
        ASSERT_EQ(4, paragraph->GetChildNodes(NodeType::Any, true)->get_Count());

        // We can modify the contents of the run by editing and deleting existing child nodes.
        (System::ExplicitCast<Run>(paragraph->GetChildNodes(NodeType::Run, true)->idx_get(1)))->set_Text(u"Updated run 2. ");
        paragraph->GetChildNodes(NodeType::Run, true)->Remove(paragraphText);

        ASSERT_EQ(u"Run 1. Updated run 2. Run 3.", paragraph->GetText().Trim());
        ASSERT_EQ(3, paragraph->GetChildNodes(NodeType::Any, true)->get_Count());
        //ExEnd
    }

    void Revisions()
    {
        //ExStart
        //ExFor:Paragraph.IsMoveFromRevision
        //ExFor:Paragraph.IsMoveToRevision
        //ExFor:ParagraphCollection
        //ExFor:ParagraphCollection.Item(Int32)
        //ExFor:Story.Paragraphs
        //ExSummary:Shows how to check whether a paragraph is a move revision.
        auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

        // This document contains "Move" revisions, which appear when we highlight text with the cursor,
        // and then drag it to move it to another location
        // while tracking revisions in Microsoft Word via "Review" -> "Track changes".
        ASSERT_EQ(6, doc->get_Revisions()->LINQ_Count([](SharedPtr<Revision> r) { return r->get_RevisionType() == RevisionType::Moving; }));

        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        // Move revisions consist of pairs of "Move from", and "Move to" revisions.
        // These revisions are potential changes to the document that we can either accept or reject.
        // Before we accept/reject a move revision, the document
        // must keep track of both the departure and arrival destinations of the text.
        // The second and the fourth paragraph define one such revision, and thus both have the same contents.
        ASSERT_EQ(paragraphs->idx_get(1)->GetText(), paragraphs->idx_get(3)->GetText());

        // The "Move from" revision is the paragraph where we dragged the text from.
        // If we accept the revision, this paragraph will disappear,
        // and the other will remain and no longer be a revision.
        ASSERT_TRUE(paragraphs->idx_get(1)->get_IsMoveFromRevision());

        // The "Move to" revision is the paragraph where we dragged the text to.
        // If we reject the revision, this paragraph instead will disappear, and the other will remain.
        ASSERT_TRUE(paragraphs->idx_get(3)->get_IsMoveToRevision());
        //ExEnd
    }

    void GetFormatRevision()
    {
        //ExStart
        //ExFor:Paragraph.IsFormatRevision
        //ExSummary:Shows how to check whether a paragraph is a format revision.
        auto doc = MakeObject<Document>(MyDir + u"Format revision.docx");

        // This paragraph is a "Format" revision, which occurs when we change the formatting of existing text
        // while tracking revisions in Microsoft Word via "Review" -> "Track changes".
        ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_IsFormatRevision());
        //ExEnd
    }

    void GetFrameProperties()
    {
        //ExStart
        //ExFor:Paragraph.FrameFormat
        //ExFor:FrameFormat
        //ExFor:FrameFormat.IsFrame
        //ExFor:FrameFormat.Width
        //ExFor:FrameFormat.Height
        //ExFor:FrameFormat.HeightRule
        //ExFor:FrameFormat.HorizontalAlignment
        //ExFor:FrameFormat.VerticalAlignment
        //ExFor:FrameFormat.HorizontalPosition
        //ExFor:FrameFormat.RelativeHorizontalPosition
        //ExFor:FrameFormat.HorizontalDistanceFromText
        //ExFor:FrameFormat.VerticalPosition
        //ExFor:FrameFormat.RelativeVerticalPosition
        //ExFor:FrameFormat.VerticalDistanceFromText
        //ExSummary:Shows how to get information about formatting properties of paragraphs that are frames.
        auto doc = MakeObject<Document>(MyDir + u"Paragraph frame.docx");

        SharedPtr<Paragraph> paragraphFrame = doc->get_FirstSection()->get_Body()->get_Paragraphs()->LINQ_OfType<SharedPtr<Paragraph>>()->LINQ_First(
            [](SharedPtr<Paragraph> p) { return p->get_FrameFormat()->get_IsFrame(); });

        ASPOSE_ASSERT_EQ(233.3, paragraphFrame->get_FrameFormat()->get_Width());
        ASPOSE_ASSERT_EQ(138.8, paragraphFrame->get_FrameFormat()->get_Height());
        ASSERT_EQ(HeightRule::AtLeast, paragraphFrame->get_FrameFormat()->get_HeightRule());
        ASSERT_EQ(HorizontalAlignment::Default, paragraphFrame->get_FrameFormat()->get_HorizontalAlignment());
        ASSERT_EQ(VerticalAlignment::Default, paragraphFrame->get_FrameFormat()->get_VerticalAlignment());
        ASPOSE_ASSERT_EQ(34.05, paragraphFrame->get_FrameFormat()->get_HorizontalPosition());
        ASSERT_EQ(RelativeHorizontalPosition::Page, paragraphFrame->get_FrameFormat()->get_RelativeHorizontalPosition());
        ASPOSE_ASSERT_EQ(9.0, paragraphFrame->get_FrameFormat()->get_HorizontalDistanceFromText());
        ASPOSE_ASSERT_EQ(20.5, paragraphFrame->get_FrameFormat()->get_VerticalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Paragraph, paragraphFrame->get_FrameFormat()->get_RelativeVerticalPosition());
        ASPOSE_ASSERT_EQ(0.0, paragraphFrame->get_FrameFormat()->get_VerticalDistanceFromText());
        //ExEnd
    }

    /// <summary>
    /// Insert field into the first paragraph of the current document using field type.
    /// </summary>
    static void InsertFieldUsingFieldType(SharedPtr<Document> doc, FieldType fieldType, bool updateField, SharedPtr<Node> refNode, bool isAfter, int paraIndex)
    {
        SharedPtr<Paragraph> para = DocumentHelper::GetParagraph(doc, paraIndex);
        para->InsertField(fieldType, updateField, refNode, isAfter);
    }

    /// <summary>
    /// Insert field into the first paragraph of the current document using field code.
    /// </summary>
    static void InsertFieldUsingFieldCode(SharedPtr<Document> doc, String fieldCode, SharedPtr<Node> refNode, bool isAfter, int paraIndex)
    {
        SharedPtr<Paragraph> para = DocumentHelper::GetParagraph(doc, paraIndex);
        para->InsertField(fieldCode, refNode, isAfter);
    }

    /// <summary>
    /// Insert field into the first paragraph of the current document using field code and field String.
    /// </summary>
    static void InsertFieldUsingFieldCodeFieldString(SharedPtr<Document> doc, String fieldCode, String fieldValue, SharedPtr<Node> refNode, bool isAfter,
                                                     int paraIndex)
    {
        SharedPtr<Paragraph> para = DocumentHelper::GetParagraph(doc, paraIndex);
        para->InsertField(fieldCode, fieldValue, refNode, isAfter);
    }

    void IsRevision()
    {
        //ExStart
        //ExFor:Paragraph.IsDeleteRevision
        //ExFor:Paragraph.IsInsertRevision
        //ExSummary:Shows how to work with revision paragraphs.
        auto doc = MakeObject<Document>();
        SharedPtr<Body> body = doc->get_FirstSection()->get_Body();
        SharedPtr<Paragraph> para = body->get_FirstParagraph();

        para->AppendChild(MakeObject<Run>(doc, u"Paragraph 1. "));
        body->AppendParagraph(u"Paragraph 2. ");
        body->AppendParagraph(u"Paragraph 3. ");

        // The above paragraphs are not revisions.
        // Paragraphs that we add after starting revision tracking will register as "Insert" revisions.
        doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());

        para = body->AppendParagraph(u"Paragraph 4. ");

        ASSERT_TRUE(para->get_IsInsertRevision());

        // Paragraphs that we remove after starting revision tracking will register as "Delete" revisions.
        SharedPtr<ParagraphCollection> paragraphs = body->get_Paragraphs();

        ASSERT_EQ(4, paragraphs->get_Count());

        para = paragraphs->idx_get(2);
        para->Remove();

        // Such paragraphs will remain until we either accept or reject the delete revision.
        // Accepting the revision will remove the paragraph for good,
        // and rejecting the revision will leave it in the document as if we never deleted it.
        ASSERT_EQ(4, paragraphs->get_Count());
        ASSERT_TRUE(para->get_IsDeleteRevision());

        // Accept the revision, and then verify that the paragraph is gone.
        doc->AcceptAllRevisions();

        ASSERT_EQ(3, paragraphs->get_Count());
        ASSERT_EQ(0, para->get_Count());
        ASSERT_EQ(String(u"Paragraph 1. \r") + u"Paragraph 2. \r" + u"Paragraph 4.", doc->GetText().Trim());
        //ExEnd
    }

    void BreakIsStyleSeparator()
    {
        //ExStart
        //ExFor:Paragraph.BreakIsStyleSeparator
        //ExSummary:Shows how to write text to the same line as a TOC heading and have it not show up in the TOC.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertTableOfContents(u"\\o \\h \\z \\u");
        builder->InsertBreak(BreakType::PageBreak);

        // Insert a paragraph with a style that the TOC will pick up as an entry.
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);

        // Both these strings are in the same paragraph and will therefore show up on the same TOC entry.
        builder->Write(u"Heading 1. ");
        builder->Write(u"Will appear in the TOC. ");

        // If we insert a style separator, we can write more text in the same paragraph
        // and use a different style without showing up in the TOC.
        // If we use a heading type style after the separator, we can draw multiple TOC entries from one document text line.
        builder->InsertStyleSeparator();
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Quote);
        builder->Write(u"Won't appear in the TOC. ");

        ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_BreakIsStyleSeparator());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Paragraph.BreakIsStyleSeparator.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Paragraph.BreakIsStyleSeparator.docx");

        TestUtil::VerifyField(
            FieldType::FieldTOC, u"TOC \\o \\h \\z \\u",
            u"\u0013 HYPERLINK \\l \"_Toc256000000\" \u0014Heading 1. Will appear in the TOC.\t\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\u0015\r",
            doc->get_Range()->get_Fields()->idx_get(0));
        ASSERT_FALSE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_BreakIsStyleSeparator());
    }

    void TabStops()
    {
        //ExStart
        //ExFor:Paragraph.GetEffectiveTabStops
        //ExSummary:Shows how to set custom tab stops for a paragraph.
        auto doc = MakeObject<Document>();
        SharedPtr<Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();

        // If we are in a paragraph with no tab stops in this collection,
        // the cursor will jump 36 points each time we press the Tab key in Microsoft Word.
        ASSERT_EQ(0, doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetEffectiveTabStops()->get_Length());

        // We can add custom tab stops in Microsoft Word if we enable the ruler via the "View" tab.
        // Each unit on this ruler is two default tab stops, which is 72 points.
        // We can add custom tab stops programmatically like this.
        SharedPtr<TabStopCollection> tabStops = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_TabStops();
        tabStops->Add(72, TabAlignment::Left, TabLeader::Dots);
        tabStops->Add(216, TabAlignment::Center, TabLeader::Dashes);
        tabStops->Add(360, TabAlignment::Right, TabLeader::Line);

        // We can see these tab stops in Microsoft Word by enabling the ruler via "View" -> "Show" -> "Ruler".
        ASSERT_EQ(3, para->GetEffectiveTabStops()->get_Length());

        // Any tab characters we add will make use of the tab stops on the ruler and may,
        // depending on the tab leader's value, leave a line between the tab departure and arrival destinations.
        para->AppendChild(MakeObject<Run>(doc, u"\tTab 1\tTab 2\tTab 3"));

        doc->Save(ArtifactsDir + u"Paragraph.TabStops.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Paragraph.TabStops.docx");
        tabStops = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_TabStops();

        TestUtil::VerifyTabStop(72.0, TabAlignment::Left, TabLeader::Dots, false, tabStops->idx_get(0));
        TestUtil::VerifyTabStop(216.0, TabAlignment::Center, TabLeader::Dashes, false, tabStops->idx_get(1));
        TestUtil::VerifyTabStop(360.0, TabAlignment::Right, TabLeader::Line, false, tabStops->idx_get(2));
    }

    void JoinRuns()
    {
        //ExStart
        //ExFor:Paragraph.JoinRunsWithSameFormatting
        //ExSummary:Shows how to simplify paragraphs by merging superfluous runs.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert four runs of text into the paragraph.
        builder->Write(u"Run 1. ");
        builder->Write(u"Run 2. ");
        builder->Write(u"Run 3. ");
        builder->Write(u"Run 4. ");

        // If we open this document in Microsoft Word, the paragraph will look like one seamless text body.
        // However, it will consist of four separate runs with the same formatting. Fragmented paragraphs like this
        // may occur when we manually edit parts of one paragraph many times in Microsoft Word.
        SharedPtr<Paragraph> para = builder->get_CurrentParagraph();

        ASSERT_EQ(4, para->get_Runs()->get_Count());

        // Change the style of the last run to set it apart from the first three.
        para->get_Runs()->idx_get(3)->get_Font()->set_StyleIdentifier(StyleIdentifier::Emphasis);

        // We can run the "JoinRunsWithSameFormatting" method to optimize the document's contents
        // by merging similar runs into one, reducing their overall count.
        // This method also returns the number of runs that this method merged.
        // These two merges occurred to combine Runs #1, #2, and #3,
        // while leaving out Run #4 because it has an incompatible style.
        ASSERT_EQ(2, para->JoinRunsWithSameFormatting());

        // The number of runs left will equal the original count
        // minus the number of run merges that the "JoinRunsWithSameFormatting" method carried out.
        ASSERT_EQ(2, para->get_Runs()->get_Count());
        ASSERT_EQ(u"Run 1. Run 2. Run 3. ", para->get_Runs()->idx_get(0)->get_Text());
        ASSERT_EQ(u"Run 4. ", para->get_Runs()->idx_get(1)->get_Text());
        //ExEnd
    }
};

} // namespace ApiExamples
