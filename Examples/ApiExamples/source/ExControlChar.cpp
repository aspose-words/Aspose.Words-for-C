// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExControlChar.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
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
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(717781891u, ::Aspose::Words::ApiExamples::ExControlChar, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExControlChar : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExControlChar> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExControlChar>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExControlChar> ExControlChar::s_instance;

} // namespace gtest_test

void ExControlChar::CarriageReturn()
{
    //ExStart
    //ExFor:ControlChar
    //ExFor:ControlChar.Cr
    //ExFor:Node.GetText
    //ExSummary:Shows how to use control characters.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert paragraphs with text with DocumentBuilder.
    builder->Writeln(u"Hello world!");
    builder->Writeln(u"Hello again!");
    
    // Converting the document to text form reveals that control characters
    // represent some of the document's structural elements, such as page breaks.
    ASSERT_EQ(System::String::Format(u"Hello world!{0}", Aspose::Words::ControlChar::Cr()) + System::String::Format(u"Hello again!{0}", Aspose::Words::ControlChar::Cr()) + Aspose::Words::ControlChar::PageBreak(), doc->GetText());
    
    // When converting a document to string form,
    // we can omit some of the control characters with the Trim method.
    ASSERT_EQ(System::String::Format(u"Hello world!{0}", Aspose::Words::ControlChar::Cr()) + u"Hello again!", doc->GetText().Trim());
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
    //ExSummary:Shows how to add various control characters to a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Add a regular space.
    builder->Write(System::String(u"Before space.") + Aspose::Words::ControlChar::SpaceChar + u"After space.");
    
    // Add an NBSP, which is a non-breaking space.
    // Unlike the regular space, this space cannot have an automatic line break at its position.
    builder->Write(System::String(u"Before space.") + Aspose::Words::ControlChar::NonBreakingSpace() + u"After space.");
    
    // Add a tab character.
    builder->Write(System::String(u"Before tab.") + Aspose::Words::ControlChar::Tab() + u"After tab.");
    
    // Add a line break.
    builder->Write(System::String(u"Before line break.") + Aspose::Words::ControlChar::LineBreak() + u"After line break.");
    
    // Add a new line and starts a new paragraph.
    ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)->get_Count());
    builder->Write(System::String(u"Before line feed.") + Aspose::Words::ControlChar::LineFeed() + u"After line feed.");
    ASSERT_EQ(2, doc->get_FirstSection()->get_Body()->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)->get_Count());
    
    // The line feed character has two versions.
    ASSERT_EQ(Aspose::Words::ControlChar::LineFeed(), Aspose::Words::ControlChar::Lf());
    
    // Carriage returns and line feeds can be represented together by one character.
    ASSERT_EQ(Aspose::Words::ControlChar::CrLf(), Aspose::Words::ControlChar::Cr() + Aspose::Words::ControlChar::Lf());
    
    // Add a paragraph break, which will start a new paragraph.
    builder->Write(System::String(u"Before paragraph break.") + Aspose::Words::ControlChar::ParagraphBreak() + u"After paragraph break.");
    ASSERT_EQ(3, doc->get_FirstSection()->get_Body()->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)->get_Count());
    
    // Add a section break. This does not make a new section or paragraph.
    ASSERT_EQ(1, doc->get_Sections()->get_Count());
    builder->Write(System::String(u"Before section break.") + Aspose::Words::ControlChar::SectionBreak() + u"After section break.");
    ASSERT_EQ(1, doc->get_Sections()->get_Count());
    
    // Add a page break.
    builder->Write(System::String(u"Before page break.") + Aspose::Words::ControlChar::PageBreak() + u"After page break.");
    
    // A page break is the same value as a section break.
    ASSERT_EQ(Aspose::Words::ControlChar::PageBreak(), Aspose::Words::ControlChar::SectionBreak());
    
    // Insert a new section, and then set its column count to two.
    doc->AppendChild<System::SharedPtr<Aspose::Words::Section>>(System::MakeObject<Aspose::Words::Section>(doc));
    builder->MoveToSection(1);
    builder->get_CurrentSection()->get_PageSetup()->get_TextColumns()->SetCount(2);
    
    // We can use a control character to mark the point where text moves to the next column.
    builder->Write(System::String(u"Text at end of column 1.") + Aspose::Words::ControlChar::ColumnBreak() + u"Text at beginning of column 2.");
    
    doc->Save(get_ArtifactsDir() + u"ControlChar.InsertControlChars.docx");
    
    // There are char and string counterparts for most characters.
    ASPOSE_ASSERT_EQ(System::Convert::ToChar(Aspose::Words::ControlChar::Cell()), Aspose::Words::ControlChar::CellChar);
    ASPOSE_ASSERT_EQ(System::Convert::ToChar(Aspose::Words::ControlChar::NonBreakingSpace()), Aspose::Words::ControlChar::NonBreakingSpaceChar);
    ASPOSE_ASSERT_EQ(System::Convert::ToChar(Aspose::Words::ControlChar::Tab()), Aspose::Words::ControlChar::TabChar);
    ASPOSE_ASSERT_EQ(System::Convert::ToChar(Aspose::Words::ControlChar::LineBreak()), Aspose::Words::ControlChar::LineBreakChar);
    ASPOSE_ASSERT_EQ(System::Convert::ToChar(Aspose::Words::ControlChar::LineFeed()), Aspose::Words::ControlChar::LineFeedChar);
    ASPOSE_ASSERT_EQ(System::Convert::ToChar(Aspose::Words::ControlChar::ParagraphBreak()), Aspose::Words::ControlChar::ParagraphBreakChar);
    ASPOSE_ASSERT_EQ(System::Convert::ToChar(Aspose::Words::ControlChar::SectionBreak()), Aspose::Words::ControlChar::SectionBreakChar);
    ASPOSE_ASSERT_EQ(System::Convert::ToChar(Aspose::Words::ControlChar::PageBreak()), Aspose::Words::ControlChar::SectionBreakChar);
    ASPOSE_ASSERT_EQ(System::Convert::ToChar(Aspose::Words::ControlChar::ColumnBreak()), Aspose::Words::ControlChar::ColumnBreakChar);
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
} // namespace Words
} // namespace Aspose
