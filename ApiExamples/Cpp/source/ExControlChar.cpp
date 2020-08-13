// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExControlChar.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/convert.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Sections/TextColumnCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
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

class ExControlChar : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExControlChar> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExControlChar>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExControlChar> ExControlChar::s_instance;

} // namespace gtest_test

void ExControlChar::CarriageReturn()
{
    //ExStart
    //ExFor:ControlChar
    //ExFor:ControlChar.Cr
    //ExSummary:Shows how to use control characters.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert paragraphs with text with DocumentBuilder
    builder->Writeln(u"Hello world!");
    builder->Writeln(u"Hello again!");

    // The entire document, when in string form, will display some structural features such as breaks with control characters
    ASSERT_EQ(String::Format(u"Hello world!{0}Hello again!{1}{2}",ControlChar::Cr(),ControlChar::Cr(),ControlChar::PageBreak()), doc->GetText());

    // Some of them can be trimmed out
    ASSERT_EQ(String::Format(u"Hello world!{0}Hello again!",ControlChar::Cr()), doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExControlChar, CarriageReturn)
{
    s_instance->CarriageReturn();
}

} // namespace gtest_test

void ExControlChar::InsertControlChars()
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
    //ExSummary:Shows how to use various control characters.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Add a regular space
    builder->Write(String(u"Before space.") + ControlChar::SpaceChar + u"After space.");

    // Add a NBSP, or non-breaking space
    // Unlike the regular space, this space can't have an automatic line break at its position
    builder->Write(String(u"Before space.") + ControlChar::NonBreakingSpace() + u"After space.");

    // Add a tab character
    builder->Write(String(u"Before tab.") + ControlChar::Tab() + u"After tab.");

    // Add a line break
    builder->Write(String(u"Before line break.") + ControlChar::LineBreak() + u"After line break.");

    // This adds a new line and starts a new paragraph
    // Same value as ControlChar.Lf
    ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)->get_Count());
    builder->Write(String(u"Before line feed.") + ControlChar::LineFeed() + u"After line feed.");
    ASSERT_EQ(2, doc->get_FirstSection()->get_Body()->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)->get_Count());

    // Carriage returns and line feeds can be represented together by one character
    ASSERT_EQ(ControlChar::CrLf(), ControlChar::Cr() + ControlChar::Lf());

    // The line feed character has two versions
    ASSERT_EQ(ControlChar::LineFeed(), ControlChar::Lf());

    // Add a paragraph break, also adding a new paragraph
    builder->Write(String(u"Before paragraph break.") + ControlChar::ParagraphBreak() + u"After paragraph break.");
    ASSERT_EQ(3, doc->get_FirstSection()->get_Body()->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)->get_Count());

    // Add a section break. Note that this does not make a new section or paragraph
    ASSERT_EQ(1, doc->get_Sections()->get_Count());
    builder->Write(String(u"Before section break.") + ControlChar::SectionBreak() + u"After section break.");
    ASSERT_EQ(1, doc->get_Sections()->get_Count());

    // A page break is the same value as a section break
    builder->Write(String(u"Before page break.") + ControlChar::PageBreak() + u"After page break.");

    // We can add a new section like this
    doc->AppendChild(MakeObject<Section>(doc));
    builder->MoveToSection(1);

    // If you have a section with more than one column, you can use a column break to make following text start on a new column
    builder->get_CurrentSection()->get_PageSetup()->get_TextColumns()->SetCount(2);
    builder->Write(String(u"Text at end of column 1.") + ControlChar::ColumnBreak() + u"Text at beginning of column 2.");

    // Save document to see the characters we added
    doc->Save(ArtifactsDir + u"ControlChar.InsertControlChars.docx");

    // There are char and string counterparts for most characters
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

namespace gtest_test
{

TEST_F(ExControlChar, InsertControlChars)
{
    s_instance->InsertControlChars();
}

} // namespace gtest_test

} // namespace ApiExamples
