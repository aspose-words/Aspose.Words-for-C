#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/EndnoteOptions.h>
#include <Aspose.Words.Cpp/Model/Footnotes/EndnotePosition.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteNumberingRule.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteOptions.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnotePosition.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Numbering/NumberStyle.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/StoryType.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <drawing/color.h>
#include <system/collections/ienumerable.h>
#include <system/collections/list.h>
#include <system/date_time.h>
#include <system/exceptions.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Tables;

namespace ApiExamples {

class ExInlineStory : public ApiExampleBase
{
public:
    void Footnotes()
    {
        //ExStart
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.NumberStyle
        //ExFor:FootnoteOptions.Position
        //ExFor:FootnoteOptions.RestartRule
        //ExFor:FootnoteOptions.StartNumber
        //ExFor:FootnoteNumberingRule
        //ExFor:FootnotePosition
        //ExSummary:Shows how to insert footnotes, and modify their appearance.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Text 1. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Write(u"Text 2. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 2");
        builder->Write(u"Text 3. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 3", u"Custom reference mark");

        doc->get_FootnoteOptions()->set_Position(FootnotePosition::BeneathText);
        doc->get_FootnoteOptions()->set_NumberStyle(NumberStyle::UppercaseRoman);
        doc->get_FootnoteOptions()->set_RestartRule(FootnoteNumberingRule::Continuous);
        doc->get_FootnoteOptions()->set_StartNumber(1);

        doc->Save(ArtifactsDir + u"InlineStory.Footnotes.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"InlineStory.Footnotes.docx");

        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote 1",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote 2",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 1, true)));
        TestUtil::VerifyFootnote(FootnoteType::Footnote, false, u"Custom reference mark", u"Custom reference mark Footnote 3",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 2, true)));
    }

    void Endnotes()
    {
        //ExStart
        //ExFor:Document.EndnoteOptions
        //ExFor:EndnoteOptions
        //ExFor:EndnoteOptions.NumberStyle
        //ExFor:EndnoteOptions.Position
        //ExFor:EndnoteOptions.RestartRule
        //ExFor:EndnoteOptions.StartNumber
        //ExFor:EndnotePosition
        //ExSummary:Shows how to insert endnotes, and modify their appearance.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Text 1. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 1");
        builder->Write(u"Text 2. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 2");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Write(u"Text 3. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 3", u"Custom reference mark");

        doc->get_EndnoteOptions()->set_Position(EndnotePosition::EndOfDocument);
        doc->get_EndnoteOptions()->set_NumberStyle(NumberStyle::UppercaseRoman);
        doc->get_EndnoteOptions()->set_RestartRule(FootnoteNumberingRule::Continuous);
        doc->get_EndnoteOptions()->set_StartNumber(1);

        doc->Save(ArtifactsDir + u"InlineStory.Endnotes.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"InlineStory.Endnotes.docx");

        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote 1",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote 2",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 1, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, false, u"Custom reference mark", u"Custom reference mark Endnote 3",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 2, true)));
    }

    void AddFootnote()
    {
        //ExStart
        //ExFor:Footnote
        //ExFor:Footnote.IsAuto
        //ExFor:Footnote.ReferenceMark
        //ExFor:InlineStory
        //ExFor:InlineStory.Paragraphs
        //ExFor:InlineStory.FirstParagraph
        //ExFor:FootnoteType
        //ExFor:Footnote.#ctor
        //ExSummary:Shows how to add a footnote to a paragraph in the document and set its marker.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add text that will be referenced by a footnote
        builder->Write(u"Main body text.");

        // Add a footnote and give it text, which will appear at the bottom of the page
        SharedPtr<Footnote> footnote = builder->InsertFootnote(FootnoteType::Footnote, u"Footnote text.");

        // This attribute means the footnote in the main text will automatically be assigned a number, "1" in this instance
        // The next footnote will get "2"
        ASSERT_TRUE(footnote->get_IsAuto());

        // We can edit the footnote's text like this
        // Make sure to move the builder back into the document body afterwards
        builder->MoveTo(footnote->get_FirstParagraph());
        builder->Write(u" More text added by a DocumentBuilder.");
        builder->MoveToDocumentEnd();

        ASSERT_EQ(u"Footnote text. More text added by a DocumentBuilder.", footnote->get_Paragraphs()->idx_get(0)->ToString(SaveFormat::Text).Trim());

        builder->Write(u" More main body text.");
        footnote = builder->InsertFootnote(FootnoteType::Footnote, u"Footnote text.");

        // Substitute the reference number for our own custom mark by setting this variable, "IsAuto" will also be set to false
        footnote->set_ReferenceMark(u"RefMark");
        ASSERT_FALSE(footnote->get_IsAuto());

        // This bookmark will get a number "3" even though there was no "2"
        builder->Write(u" More main body text.");
        footnote = builder->InsertFootnote(FootnoteType::Footnote, u"Footnote text.");
        ASSERT_TRUE(footnote->get_IsAuto());

        doc->Save(ArtifactsDir + u"InlineStory.AddFootnote.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"InlineStory.AddFootnote.docx");

        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote text. More text added by a DocumentBuilder.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
        TestUtil::VerifyFootnote(FootnoteType::Footnote, false, u"RefMark", u"Footnote text.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 1, true)));
        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote text.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 2, true)));
    }

    void FootnoteEndnote()
    {
        //ExStart
        //ExFor:Footnote.FootnoteType
        //ExSummary:Demonstrates the difference between footnotes and endnotes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Write text and insert a footnote to reference it at the bottom of the page
        builder->Write(u"Footnote referenced main body text.");
        SharedPtr<Footnote> footnote =
            builder->InsertFootnote(FootnoteType::Footnote, u"Footnote text, will appear at the bottom of the page that contains the referenced text.");

        // Write text and insert an endnote to reference it at the end of the document
        builder->Write(u"Endnote referenced main body text.");
        SharedPtr<Footnote> endnote = builder->InsertFootnote(FootnoteType::Endnote, u"Endnote text, will appear at the very end of the document.");

        // Since endnotes are at the end of the document, breaks like this will push them down while the footnotes stay where they are
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->InsertBreak(BreakType::SectionBreakNewPage);

        ASSERT_EQ(FootnoteType::Footnote, footnote->get_FootnoteType());
        ASSERT_EQ(FootnoteType::Endnote, endnote->get_FootnoteType());

        doc->Save(ArtifactsDir + u"InlineStory.FootnoteEndnote.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"InlineStory.FootnoteEndnote.docx");

        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty,
                                 u"Footnote text, will appear at the bottom of the page that contains the referenced text.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote text, will appear at the very end of the document.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 1, true)));
    }

    void AddComment()
    {
        //ExStart
        //ExFor:Comment
        //ExFor:InlineStory
        //ExFor:InlineStory.Paragraphs
        //ExFor:InlineStory.FirstParagraph
        //ExFor:Comment.#ctor(DocumentBase, String, String, DateTime)
        //ExSummary:Shows how to add a comment to a paragraph in the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Some text is added.");

        auto comment = MakeObject<Comment>(doc, u"Amy Lee", u"AL", System::DateTime::get_Today());
        builder->get_CurrentParagraph()->AppendChild(comment);
        comment->get_Paragraphs()->Add(MakeObject<Paragraph>(doc));
        comment->get_FirstParagraph()->get_Runs()->Add(MakeObject<Run>(doc, u"Comment text."));

        doc->Save(ArtifactsDir + u"InlineStory.AddComment.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"InlineStory.AddComment.docx");
        comment = System::DynamicCast<Comment>(doc->GetChild(NodeType::Comment, 0, true));

        ASSERT_EQ(u"Comment text.\r", comment->GetText());
        ASSERT_EQ(u"Amy Lee", comment->get_Author());
        ASSERT_EQ(u"AL", comment->get_Initial());
        ASSERT_EQ(System::DateTime::get_Today(), comment->get_DateTime());
    }

    void InlineStoryRevisions()
    {
        //ExStart
        //ExFor:InlineStory.IsDeleteRevision
        //ExFor:InlineStory.IsInsertRevision
        //ExFor:InlineStory.IsMoveFromRevision
        //ExFor:InlineStory.IsMoveToRevision
        //ExSummary:Shows how to view revision-related properties of InlineStory nodes.
        // Open a document that has revisions from changes being tracked
        auto doc = MakeObject<Document>(MyDir + u"Revision footnotes.docx");
        ASSERT_TRUE(doc->get_HasRevisions());

        // Get a collection of all footnotes from the document
        SharedPtr<System::Collections::Generic::List<SharedPtr<Footnote>>> footnotes =
            doc->GetChildNodes(NodeType::Footnote, true)->LINQ_Cast<SharedPtr<Footnote>>()->LINQ_ToList();
        ASSERT_EQ(5, footnotes->get_Count());

        // If a node was inserted in Microsoft Word while changes were being tracked, this flag will be set to true
        ASSERT_TRUE(footnotes->idx_get(2)->get_IsInsertRevision());

        // If one node was moved from one place to another while changes were tracked,
        // the node will be placed at the departure location as a "move to revision",
        // and a "move from revision" node will be left behind at the origin, in case we want to reject changes
        // Highlighting text and dragging it to another place with the mouse and cut-and-pasting (but not copy-pasting) both count as "move revisions"
        // The node with the "IsMoveToRevision" flag is the arrival of the move operation, and the node with the "IsMoveFromRevision" flag is the departure
        // point
        ASSERT_TRUE(footnotes->idx_get(1)->get_IsMoveToRevision());
        ASSERT_TRUE(footnotes->idx_get(4)->get_IsMoveFromRevision());

        // If a node was deleted while changes were being tracked, it will stay behind as a delete revision until we accept/reject changes
        ASSERT_TRUE(footnotes->idx_get(3)->get_IsDeleteRevision());
        //ExEnd
    }

    void InsertInlineStoryNodes()
    {
        //ExStart
        //ExFor:Comment.StoryType
        //ExFor:Footnote.StoryType
        //ExFor:InlineStory.EnsureMinimum
        //ExFor:InlineStory.Font
        //ExFor:InlineStory.LastParagraph
        //ExFor:InlineStory.ParentParagraph
        //ExFor:InlineStory.StoryType
        //ExFor:InlineStory.Tables
        //ExSummary:Shows how to insert InlineStory nodes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        SharedPtr<Footnote> footnote = builder->InsertFootnote(FootnoteType::Footnote, nullptr);

        // Table nodes have an "EnsureMinimum()" method that makes sure the table has at least one cell
        auto table = MakeObject<Table>(doc);
        table->EnsureMinimum();

        // We can place a table inside a footnote, which will make it appear at the footer of the referencing page
        ASSERT_EQ(0, footnote->get_Tables()->get_Count());
        footnote->AppendChild(table);
        ASSERT_EQ(1, footnote->get_Tables()->get_Count());
        ASSERT_EQ(NodeType::Table, footnote->get_LastChild()->get_NodeType());

        // An InlineStory has an "EnsureMinimum()" method as well, but in this case,
        // it makes sure the last child of the node is a paragraph, in order for us to be able to click and write text easily in Microsoft Word
        footnote->EnsureMinimum();
        ASSERT_EQ(NodeType::Paragraph, footnote->get_LastChild()->get_NodeType());

        // Edit the appearance of the anchor, which is the small superscript number in the main text that points to the footnote
        footnote->get_Font()->set_Name(u"Arial");
        footnote->get_Font()->set_Color(System::Drawing::Color::get_Green());

        // All inline story nodes have their own respective story types
        ASSERT_EQ(StoryType::Footnotes, footnote->get_StoryType());

        // A comment is another type of inline story
        auto comment = System::DynamicCast<Comment>(
            builder->get_CurrentParagraph()->AppendChild(MakeObject<Comment>(doc, u"John Doe", u"J. D.", System::DateTime::get_Now())));

        // The parent paragraph of an inline story node will be the one from the main document body
        ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_FirstParagraph(), comment->get_ParentParagraph());

        // However, the last paragraph is the one from the comment text contents, which will be outside the main document body in a speech bubble
        // A comment won't have any child nodes by default, so we can apply the EnsureMinimum() method to place a paragraph here as well
        ASSERT_TRUE(comment->get_LastParagraph() == nullptr);
        comment->EnsureMinimum();
        ASSERT_EQ(NodeType::Paragraph, comment->get_LastChild()->get_NodeType());

        // Once we have a paragraph, we can move the builder do it and write our comment
        builder->MoveTo(comment->get_LastParagraph());
        builder->Write(u"My comment.");

        ASSERT_EQ(StoryType::Comments, comment->get_StoryType());

        doc->Save(ArtifactsDir + u"InlineStory.InsertInlineStoryNodes.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"InlineStory.InsertInlineStoryNodes.docx");

        footnote = System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true));

        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, String::Empty,
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
        ASSERT_EQ(u"Arial", footnote->get_Font()->get_Name());
        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), footnote->get_Font()->get_Color().ToArgb());

        comment = System::DynamicCast<Comment>(doc->GetChild(NodeType::Comment, 0, true));

        ASSERT_EQ(u"My comment.", comment->ToString(SaveFormat::Text).Trim());
    }

    void DeleteShapes()
    {
        //ExStart
        //ExFor:Story
        //ExFor:Story.DeleteShapes
        //ExFor:Story.StoryType
        //ExFor:StoryType
        //ExSummary:Shows how to clear a body of inline shapes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a DocumentBuilder to insert a shape
        // This is an inline shape, which has a parent Paragraph, which is in turn a child of the Body
        builder->InsertShape(ShapeType::Cube, 100.0, 100.0);

        ASSERT_EQ(1, doc->GetChildNodes(NodeType::Shape, true)->get_Count());

        // We can delete all such shapes from the Body, affecting all child Paragraphs
        ASSERT_EQ(StoryType::MainText, doc->get_FirstSection()->get_Body()->get_StoryType());
        doc->get_FirstSection()->get_Body()->DeleteShapes();

        ASSERT_EQ(0, doc->GetChildNodes(NodeType::Shape, true)->get_Count());
        //ExEnd
    }
};

} // namespace ApiExamples
