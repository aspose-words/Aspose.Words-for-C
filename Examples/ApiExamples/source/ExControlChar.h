#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/ControlChar.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/TextColumnCollection.h>
#include <system/convert.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace ApiExamples {

class ExControlChar : public ApiExampleBase
{
public:
    void CarriageReturn()
    {
        //ExStart
        //ExFor:ControlChar
        //ExFor:ControlChar.Cr
        //ExFor:Node.GetText
        //ExSummary:Shows how to use control characters.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert paragraphs with text with DocumentBuilder.
        builder->Writeln(u"Hello world!");
        builder->Writeln(u"Hello again!");

        // Converting the document to text form reveals that control characters
        // represent some of the document's structural elements, such as page breaks.
        ASSERT_EQ(String::Format(u"Hello world!{0}", ControlChar::Cr()) + String::Format(u"Hello again!{0}", ControlChar::Cr()) + ControlChar::PageBreak(),
                  doc->GetText());

        // When converting a document to string form,
        // we can omit some of the control characters with the Trim method.
        ASSERT_EQ(String::Format(u"Hello world!{0}", ControlChar::Cr()) + u"Hello again!", doc->GetText().Trim());
        //ExEnd
    }

    void InsertControlChars()
    {
        //ExStart
        //ExFor:ControlChar.Cell
        //ExFor:ControlChar.ColumnBreak
        //ExFor:ControlChar.CrLf
        //ExFor:ControlChar.Lf
        //ExFor:ControlChar.LineBreak
        //ExFor:ControlChar.LineFeed
        //ExFor:ControlChar.NonBreakingSpace
        //ExFor:ControlChar.PageBreak
        //ExFor:ControlChar.ParagraphBreak
        //ExFor:ControlChar.SectionBreak
        //ExFor:ControlChar.CellChar
        //ExFor:ControlChar.ColumnBreakChar
        //ExFor:ControlChar.DefaultTextInputChar
        //ExFor:ControlChar.FieldEndChar
        //ExFor:ControlChar.FieldStartChar
        //ExFor:ControlChar.FieldSeparatorChar
        //ExFor:ControlChar.LineBreakChar
        //ExFor:ControlChar.LineFeedChar
        //ExFor:ControlChar.NonBreakingHyphenChar
        //ExFor:ControlChar.NonBreakingSpaceChar
        //ExFor:ControlChar.OptionalHyphenChar
        //ExFor:ControlChar.PageBreakChar
        //ExFor:ControlChar.ParagraphBreakChar
        //ExFor:ControlChar.SectionBreakChar
        //ExFor:ControlChar.SpaceChar
        //ExSummary:Shows how to add various control characters to a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add a regular space.
        builder->Write(String(u"Before space.") + ControlChar::SpaceChar + u"After space.");

        // Add an NBSP, which is a non-breaking space.
        // Unlike the regular space, this space cannot have an automatic line break at its position.
        builder->Write(String(u"Before space.") + ControlChar::NonBreakingSpace() + u"After space.");

        // Add a tab character.
        builder->Write(String(u"Before tab.") + ControlChar::Tab() + u"After tab.");

        // Add a line break.
        builder->Write(String(u"Before line break.") + ControlChar::LineBreak() + u"After line break.");

        // Add a new line and starts a new paragraph.
        ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->GetChildNodes(NodeType::Paragraph, true)->get_Count());
        builder->Write(String(u"Before line feed.") + ControlChar::LineFeed() + u"After line feed.");
        ASSERT_EQ(2, doc->get_FirstSection()->get_Body()->GetChildNodes(NodeType::Paragraph, true)->get_Count());

        // The line feed character has two versions.
        ASSERT_EQ(ControlChar::LineFeed(), ControlChar::Lf());

        // Carriage returns and line feeds can be represented together by one character.
        ASSERT_EQ(ControlChar::CrLf(), ControlChar::Cr() + ControlChar::Lf());

        // Add a paragraph break, which will start a new paragraph.
        builder->Write(String(u"Before paragraph break.") + ControlChar::ParagraphBreak() + u"After paragraph break.");
        ASSERT_EQ(3, doc->get_FirstSection()->get_Body()->GetChildNodes(NodeType::Paragraph, true)->get_Count());

        // Add a section break. This does not make a new section or paragraph.
        ASSERT_EQ(1, doc->get_Sections()->get_Count());
        builder->Write(String(u"Before section break.") + ControlChar::SectionBreak() + u"After section break.");
        ASSERT_EQ(1, doc->get_Sections()->get_Count());

        // Add a page break.
        builder->Write(String(u"Before page break.") + ControlChar::PageBreak() + u"After page break.");

        // A page break is the same value as a section break.
        ASSERT_EQ(ControlChar::PageBreak(), ControlChar::SectionBreak());

        // Insert a new section, and then set its column count to two.
        doc->AppendChild(MakeObject<Section>(doc));
        builder->MoveToSection(1);
        builder->get_CurrentSection()->get_PageSetup()->get_TextColumns()->SetCount(2);

        // We can use a control character to mark the point where text moves to the next column.
        builder->Write(String(u"Text at end of column 1.") + ControlChar::ColumnBreak() + u"Text at beginning of column 2.");

        doc->Save(ArtifactsDir + u"ControlChar.InsertControlChars.docx");

        // There are char and string counterparts for most characters.
        ASPOSE_ASSERT_EQ(System::Convert::ToChar(ControlChar::Cell()), ControlChar::CellChar);
        ASPOSE_ASSERT_EQ(System::Convert::ToChar(ControlChar::NonBreakingSpace()), ControlChar::NonBreakingSpaceChar);
        ASPOSE_ASSERT_EQ(System::Convert::ToChar(ControlChar::Tab()), ControlChar::TabChar);
        ASPOSE_ASSERT_EQ(System::Convert::ToChar(ControlChar::LineBreak()), ControlChar::LineBreakChar);
        ASPOSE_ASSERT_EQ(System::Convert::ToChar(ControlChar::LineFeed()), ControlChar::LineFeedChar);
        ASPOSE_ASSERT_EQ(System::Convert::ToChar(ControlChar::ParagraphBreak()), ControlChar::ParagraphBreakChar);
        ASPOSE_ASSERT_EQ(System::Convert::ToChar(ControlChar::SectionBreak()), ControlChar::SectionBreakChar);
        ASPOSE_ASSERT_EQ(System::Convert::ToChar(ControlChar::PageBreak()), ControlChar::SectionBreakChar);
        ASPOSE_ASSERT_EQ(System::Convert::ToChar(ControlChar::ColumnBreak()), ControlChar::ColumnBreakChar);
        //ExEnd
    }
};

} // namespace ApiExamples
