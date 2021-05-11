#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Comment.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Notes/EndnoteOptions.h>
#include <Aspose.Words.Cpp/Notes/EndnotePosition.h>
#include <Aspose.Words.Cpp/Notes/Footnote.h>
#include <Aspose.Words.Cpp/Notes/FootnoteNumberingRule.h>
#include <Aspose.Words.Cpp/Notes/FootnoteOptions.h>
#include <Aspose.Words.Cpp/Notes/FootnotePosition.h>
#include <Aspose.Words.Cpp/Notes/FootnoteType.h>
#include <Aspose.Words.Cpp/NumberStyle.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/StoryType.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/Tables/TableCollection.h>
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

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;

namespace ApiExamples {

class ExInlineStory : public ApiExampleBase
{
public:
    void PositionFootnote(FootnotePosition footnotePosition)
    {
        //ExStart
        //ExFor:Document.FootnoteOptions
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.Position
        //ExFor:FootnotePosition
        //ExSummary:Shows how to select a different place where the document collects and displays its footnotes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // A footnote is a way to attach a reference or a side comment to text
        // that does not interfere with the main body text's flow.
        // Inserting a footnote adds a small superscript reference symbol
        // at the main body text where we insert the footnote.
        // Each footnote also creates an entry at the bottom of the page, consisting of a symbol
        // that matches the reference symbol in the main body text.
        // The reference text that we pass to the document builder's "InsertFootnote" method.
        builder->Write(u"Hello world!");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote contents.");

        // We can use the "Position" property to determine where the document will place all its footnotes.
        // If we set the value of the "Position" property to "FootnotePosition.BottomOfPage",
        // every footnote will show up at the bottom of the page that contains its reference mark. This is the default value.
        // If we set the value of the "Position" property to "FootnotePosition.BeneathText",
        // every footnote will show up at the end of the page's text that contains its reference mark.
        doc->get_FootnoteOptions()->set_Position(footnotePosition);

        doc->Save(ArtifactsDir + u"InlineStory.PositionFootnote.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"InlineStory.PositionFootnote.docx");

        ASSERT_EQ(footnotePosition, doc->get_FootnoteOptions()->get_Position());

        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote contents.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
    }

    void PositionEndnote(EndnotePosition endnotePosition)
    {
        //ExStart
        //ExFor:Document.EndnoteOptions
        //ExFor:EndnoteOptions
        //ExFor:EndnoteOptions.Position
        //ExFor:EndnotePosition
        //ExSummary:Shows how to select a different place where the document collects and displays its endnotes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // An endnote is a way to attach a reference or a side comment to text
        // that does not interfere with the main body text's flow.
        // Inserting an endnote adds a small superscript reference symbol
        // at the main body text where we insert the endnote.
        // Each endnote also creates an entry at the end of the document, consisting of a symbol
        // that matches the reference symbol in the main body text.
        // The reference text that we pass to the document builder's "InsertEndnote" method.
        builder->Write(u"Hello world!");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote contents.");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Write(u"This is the second section.");

        // We can use the "Position" property to determine where the document will place all its endnotes.
        // If we set the value of the "Position" property to "EndnotePosition.EndOfDocument",
        // every footnote will show up in a collection at the end of the document. This is the default value.
        // If we set the value of the "Position" property to "EndnotePosition.EndOfSection",
        // every footnote will show up in a collection at the end of the section whose text contains the endnote's reference mark.
        doc->get_EndnoteOptions()->set_Position(endnotePosition);

        doc->Save(ArtifactsDir + u"InlineStory.PositionEndnote.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"InlineStory.PositionEndnote.docx");

        ASSERT_EQ(endnotePosition, doc->get_EndnoteOptions()->get_Position());

        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote contents.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
    }

    void RefMarkNumberStyle()
    {
        //ExStart
        //ExFor:Document.EndnoteOptions
        //ExFor:EndnoteOptions
        //ExFor:EndnoteOptions.NumberStyle
        //ExFor:Document.FootnoteOptions
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.NumberStyle
        //ExSummary:Shows how to change the number style of footnote/endnote reference marks.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Footnotes and endnotes are a way to attach a reference or a side comment to text
        // that does not interfere with the main body text's flow.
        // Inserting a footnote/endnote adds a small superscript reference symbol
        // at the main body text where we insert the footnote/endnote.
        // Each footnote/endnote also creates an entry, which consists of a symbol that matches the reference
        // symbol in the main body text. The reference text that we pass to the document builder's "InsertEndnote" method.
        // Footnote entries, by default, show up at the bottom of each page that contains
        // their reference symbols, and endnotes show up at the end of the document.
        builder->Write(u"Text 1. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 1.");
        builder->Write(u"Text 2. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 2.");
        builder->Write(u"Text 3. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 3.", u"Custom footnote reference mark");

        builder->InsertParagraph();

        builder->Write(u"Text 1. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 1.");
        builder->Write(u"Text 2. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 2.");
        builder->Write(u"Text 3. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 3.", u"Custom endnote reference mark");

        // By default, the reference symbol for each footnote and endnote is its index
        // among all the document's footnotes/endnotes. Each document maintains separate counts
        // for footnotes and for endnotes. By default, footnotes display their numbers using Arabic numerals,
        // and endnotes display their numbers in lowercase Roman numerals.
        ASSERT_EQ(NumberStyle::Arabic, doc->get_FootnoteOptions()->get_NumberStyle());
        ASSERT_EQ(NumberStyle::LowercaseRoman, doc->get_EndnoteOptions()->get_NumberStyle());

        // We can use the "NumberStyle" property to apply custom numbering styles to footnotes and endnotes.
        // This will not affect footnotes/endnotes with custom reference marks.
        doc->get_FootnoteOptions()->set_NumberStyle(NumberStyle::UppercaseRoman);
        doc->get_EndnoteOptions()->set_NumberStyle(NumberStyle::UppercaseLetter);

        doc->Save(ArtifactsDir + u"InlineStory.RefMarkNumberStyle.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"InlineStory.RefMarkNumberStyle.docx");

        ASSERT_EQ(NumberStyle::UppercaseRoman, doc->get_FootnoteOptions()->get_NumberStyle());
        ASSERT_EQ(NumberStyle::UppercaseLetter, doc->get_EndnoteOptions()->get_NumberStyle());

        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote 1.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote 2.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 1, true)));
        TestUtil::VerifyFootnote(FootnoteType::Footnote, false, u"Custom footnote reference mark", u"Custom footnote reference mark Footnote 3.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 2, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote 1.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 3, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote 2.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 4, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, false, u"Custom endnote reference mark", u"Custom endnote reference mark Endnote 3.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 5, true)));
    }

    void NumberingRule()
    {
        //ExStart
        //ExFor:Document.EndnoteOptions
        //ExFor:EndnoteOptions
        //ExFor:EndnoteOptions.RestartRule
        //ExFor:FootnoteNumberingRule
        //ExFor:Document.FootnoteOptions
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.RestartRule
        //ExSummary:Shows how to restart footnote/endnote numbering at certain places in the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Footnotes and endnotes are a way to attach a reference or a side comment to text
        // that does not interfere with the main body text's flow.
        // Inserting a footnote/endnote adds a small superscript reference symbol
        // at the main body text where we insert the footnote/endnote.
        // Each footnote/endnote also creates an entry, which consists of a symbol that matches the reference
        // symbol in the main body text. The reference text that we pass to the document builder's "InsertEndnote" method.
        // Footnote entries, by default, show up at the bottom of each page that contains
        // their reference symbols, and endnotes show up at the end of the document.
        builder->Write(u"Text 1. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 1.");
        builder->Write(u"Text 2. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 2.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Write(u"Text 3. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 3.");
        builder->Write(u"Text 4. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 4.");

        builder->InsertBreak(BreakType::PageBreak);

        builder->Write(u"Text 1. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 1.");
        builder->Write(u"Text 2. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 2.");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Write(u"Text 3. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 3.");
        builder->Write(u"Text 4. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 4.");

        // By default, the reference symbol for each footnote and endnote is its index
        // among all the document's footnotes/endnotes. Each document maintains separate counts
        // for footnotes and endnotes and does not restart these counts at any point.
        ASSERT_EQ(doc->get_FootnoteOptions()->get_RestartRule(), FootnoteNumberingRule::Default);
        ASSERT_EQ(FootnoteNumberingRule::Default, FootnoteNumberingRule::Continuous);

        // We can use the "RestartRule" property to get the document to restart
        // the footnote/endnote counts at a new page or section.
        doc->get_FootnoteOptions()->set_RestartRule(FootnoteNumberingRule::RestartPage);
        doc->get_EndnoteOptions()->set_RestartRule(FootnoteNumberingRule::RestartSection);

        doc->Save(ArtifactsDir + u"InlineStory.NumberingRule.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"InlineStory.NumberingRule.docx");

        ASSERT_EQ(FootnoteNumberingRule::RestartPage, doc->get_FootnoteOptions()->get_RestartRule());
        ASSERT_EQ(FootnoteNumberingRule::RestartSection, doc->get_EndnoteOptions()->get_RestartRule());

        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote 1.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote 2.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 1, true)));
        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote 3.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 2, true)));
        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote 4.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 3, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote 1.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 4, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote 2.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 5, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote 3.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 6, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote 4.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 7, true)));
    }

    void StartNumber()
    {
        //ExStart
        //ExFor:Document.EndnoteOptions
        //ExFor:EndnoteOptions
        //ExFor:EndnoteOptions.StartNumber
        //ExFor:Document.FootnoteOptions
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.StartNumber
        //ExSummary:Shows how to set a number at which the document begins the footnote/endnote count.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Footnotes and endnotes are a way to attach a reference or a side comment to text
        // that does not interfere with the main body text's flow.
        // Inserting a footnote/endnote adds a small superscript reference symbol
        // at the main body text where we insert the footnote/endnote.
        // Each footnote/endnote also creates an entry, which consists of a symbol
        // that matches the reference symbol in the main body text.
        // The reference text that we pass to the document builder's "InsertEndnote" method.
        // Footnote entries, by default, show up at the bottom of each page that contains
        // their reference symbols, and endnotes show up at the end of the document.
        builder->Write(u"Text 1. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 1.");
        builder->Write(u"Text 2. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 2.");
        builder->Write(u"Text 3. ");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote 3.");

        builder->InsertParagraph();

        builder->Write(u"Text 1. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 1.");
        builder->Write(u"Text 2. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 2.");
        builder->Write(u"Text 3. ");
        builder->InsertFootnote(FootnoteType::Endnote, u"Endnote 3.");

        // By default, the reference symbol for each footnote and endnote is its index
        // among all the document's footnotes/endnotes. Each document maintains separate counts
        // for footnotes and for endnotes, which both begin at 1.
        ASSERT_EQ(1, doc->get_FootnoteOptions()->get_StartNumber());
        ASSERT_EQ(1, doc->get_EndnoteOptions()->get_StartNumber());

        // We can use the "StartNumber" property to get the document to
        // begin a footnote or endnote count at a different number.
        doc->get_EndnoteOptions()->set_NumberStyle(NumberStyle::Arabic);
        doc->get_EndnoteOptions()->set_StartNumber(50);

        doc->Save(ArtifactsDir + u"InlineStory.StartNumber.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"InlineStory.StartNumber.docx");

        ASSERT_EQ(1, doc->get_FootnoteOptions()->get_StartNumber());
        ASSERT_EQ(50, doc->get_EndnoteOptions()->get_StartNumber());
        ASSERT_EQ(NumberStyle::Arabic, doc->get_FootnoteOptions()->get_NumberStyle());
        ASSERT_EQ(NumberStyle::Arabic, doc->get_EndnoteOptions()->get_NumberStyle());

        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote 1.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote 2.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 1, true)));
        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Footnote 3.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 2, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote 1.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 3, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote 2.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 4, true)));
        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"Endnote 3.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 5, true)));
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
        //ExSummary:Shows how to insert and customize footnotes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add text, and reference it with a footnote. This footnote will place a small superscript reference
        // mark after the text that it references and create an entry below the main body text at the bottom of the page.
        // This entry will contain the footnote's reference mark and the reference text,
        // which we will pass to the document builder's "InsertFootnote" method.
        builder->Write(u"Main body text.");
        SharedPtr<Footnote> footnote = builder->InsertFootnote(FootnoteType::Footnote, u"Footnote text.");

        // If this property is set to "true", then our footnote's reference mark
        // will be its index among all the section's footnotes.
        // This is the first footnote, so the reference mark will be "1".
        ASSERT_TRUE(footnote->get_IsAuto());

        // We can move the document builder inside the footnote to edit its reference text.
        builder->MoveTo(footnote->get_FirstParagraph());
        builder->Write(u" More text added by a DocumentBuilder.");
        builder->MoveToDocumentEnd();

        ASSERT_EQ(u"\u0002 Footnote text. More text added by a DocumentBuilder.", footnote->GetText().Trim());

        builder->Write(u" More main body text.");
        footnote = builder->InsertFootnote(FootnoteType::Footnote, u"Footnote text.");

        // We can set a custom reference mark which the footnote will use instead of its index number.
        footnote->set_ReferenceMark(u"RefMark");

        ASSERT_FALSE(footnote->get_IsAuto());

        // A bookmark with the "IsAuto" flag set to true will still show its real index
        // even if previous bookmarks display custom reference marks, so this bookmark's reference mark will be a "3".
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
        //ExSummary:Shows the difference between footnotes and endnotes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are two ways of attaching numbered references to the text. Both these references will add a
        // small superscript reference mark at the location that we insert them.
        // The reference mark, by default, is the index number of the reference among all the references in the document.
        // Each reference will also create an entry, which will have the same reference mark as in the body text
        // and reference text, which we will pass to the document builder's "InsertFootnote" method.
        // 1 -  A footnote, whose entry will appear on the same page as the text that it references:
        builder->Write(u"Footnote referenced main body text.");
        SharedPtr<Footnote> footnote =
            builder->InsertFootnote(FootnoteType::Footnote, u"Footnote text, will appear at the bottom of the page that contains the referenced text.");

        // 2 -  An endnote, whose entry will appear at the end of the document:
        builder->Write(u"Endnote referenced main body text.");
        SharedPtr<Footnote> endnote = builder->InsertFootnote(FootnoteType::Endnote, u"Endnote text, will appear at the very end of the document.");

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
        //ExSummary:Shows how to add a comment to a paragraph.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Hello world!");

        auto comment = MakeObject<Comment>(doc, u"John Doe", u"JD", System::DateTime::get_Today());
        builder->get_CurrentParagraph()->AppendChild(comment);
        builder->MoveTo(comment->AppendChild(MakeObject<Paragraph>(doc)));
        builder->Write(u"Comment text.");

        ASSERT_EQ(System::DateTime::get_Today(), comment->get_DateTime());

        // In Microsoft Word, we can right-click this comment in the document body to edit it, or reply to it.
        doc->Save(ArtifactsDir + u"InlineStory.AddComment.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"InlineStory.AddComment.docx");
        comment = System::DynamicCast<Comment>(doc->GetChild(NodeType::Comment, 0, true));

        ASSERT_EQ(u"Comment text.\r", comment->GetText());
        ASSERT_EQ(u"John Doe", comment->get_Author());
        ASSERT_EQ(u"JD", comment->get_Initial());
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
        auto doc = MakeObject<Document>(MyDir + u"Revision footnotes.docx");

        // When we edit the document while the "Track Changes" option, found in via Review -> Tracking,
        // is turned on in Microsoft Word, the changes we apply count as revisions.
        // When editing a document using Aspose.Words, we can begin tracking revisions by
        // invoking the document's "StartTrackRevisions" method and stop tracking by using the "StopTrackRevisions" method.
        // We can either accept revisions to assimilate them into the document
        // or reject them to undo and discard the proposed change.
        ASSERT_TRUE(doc->get_HasRevisions());

        SharedPtr<System::Collections::Generic::List<SharedPtr<Footnote>>> footnotes =
            doc->GetChildNodes(NodeType::Footnote, true)->LINQ_Cast<SharedPtr<Footnote>>()->LINQ_ToList();

        ASSERT_EQ(5, footnotes->get_Count());

        // Below are five types of revisions that can flag an InlineStory node.
        // 1 -  An "insert" revision:
        // This revision occurs when we insert text while tracking changes.
        ASSERT_TRUE(footnotes->idx_get(2)->get_IsInsertRevision());

        // 2 -  A "move from" revision:
        // When we highlight text in Microsoft Word, and then drag it to a different place in the document
        // while tracking changes, two revisions appear.
        // The "move from" revision is a copy of the text originally before we moved it.
        ASSERT_TRUE(footnotes->idx_get(4)->get_IsMoveFromRevision());

        // 3 -  A "move to" revision:
        // The "move to" revision is the text that we moved in its new position in the document.
        // "Move from" and "move to" revisions appear in pairs for every move revision we carry out.
        // Accepting a move revision deletes the "move from" revision and its text,
        // and keeps the text from the "move to" revision.
        // Rejecting a move revision conversely keeps the "move from" revision and deletes the "move to" revision.
        ASSERT_TRUE(footnotes->idx_get(1)->get_IsMoveToRevision());

        // 4 -  A "delete" revision:
        // This revision occurs when we delete text while tracking changes. When we delete text like this,
        // it will stay in the document as a revision until we either accept the revision,
        // which will delete the text for good, or reject the revision, which will keep the text we deleted where it was.
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

        // Table nodes have an "EnsureMinimum()" method that makes sure the table has at least one cell.
        auto table = MakeObject<Table>(doc);
        table->EnsureMinimum();

        // We can place a table inside a footnote, which will make it appear at the referencing page's footer.
        ASSERT_EQ(0, footnote->get_Tables()->get_Count());
        footnote->AppendChild(table);
        ASSERT_EQ(1, footnote->get_Tables()->get_Count());
        ASSERT_EQ(NodeType::Table, footnote->get_LastChild()->get_NodeType());

        // An InlineStory has an "EnsureMinimum()" method as well, but in this case,
        // it makes sure the last child of the node is a paragraph,
        // for us to be able to click and write text easily in Microsoft Word.
        footnote->EnsureMinimum();
        ASSERT_EQ(NodeType::Paragraph, footnote->get_LastChild()->get_NodeType());

        // Edit the appearance of the anchor, which is the small superscript number
        // in the main text that points to the footnote.
        footnote->get_Font()->set_Name(u"Arial");
        footnote->get_Font()->set_Color(System::Drawing::Color::get_Green());

        // All inline story nodes have their respective story types.
        ASSERT_EQ(StoryType::Footnotes, footnote->get_StoryType());

        // A comment is another type of inline story.
        auto comment = System::DynamicCast<Comment>(
            builder->get_CurrentParagraph()->AppendChild(MakeObject<Comment>(doc, u"John Doe", u"J. D.", System::DateTime::get_Now())));

        // The parent paragraph of an inline story node will be the one from the main document body.
        ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_FirstParagraph(), comment->get_ParentParagraph());

        // However, the last paragraph is the one from the comment text contents,
        // which will be outside the main document body in a speech bubble.
        // A comment will not have any child nodes by default,
        // so we can apply the EnsureMinimum() method to place a paragraph here as well.
        ASSERT_TRUE(comment->get_LastParagraph() == nullptr);
        comment->EnsureMinimum();
        ASSERT_EQ(NodeType::Paragraph, comment->get_LastChild()->get_NodeType());

        // Once we have a paragraph, we can move the builder to do it and write our comment.
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
        //ExSummary:Shows how to remove all shapes from a node.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a DocumentBuilder to insert a shape. This is an inline shape,
        // which has a parent Paragraph, which is a child node of the first section's Body.
        builder->InsertShape(ShapeType::Cube, 100.0, 100.0);

        ASSERT_EQ(1, doc->GetChildNodes(NodeType::Shape, true)->get_Count());

        // We can delete all shapes from the child paragraphs of this Body.
        ASSERT_EQ(StoryType::MainText, doc->get_FirstSection()->get_Body()->get_StoryType());
        doc->get_FirstSection()->get_Body()->DeleteShapes();

        ASSERT_EQ(0, doc->GetChildNodes(NodeType::Shape, true)->get_Count());
        //ExEnd
    }
};

} // namespace ApiExamples
