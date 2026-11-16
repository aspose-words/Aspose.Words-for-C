#include "ExControlChar.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/convert.h>
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
    ASSERT_EQ(doc->GetText(), System::String::Format(u"Hello world!{0}", ControlChar::Cr()) + System::String::Format(u"Hello again!{0}", ControlChar::Cr()) + ControlChar::PageBreak());
    
    // When converting a document to string form,
    // we can omit some of the control characters with the Trim method.
    ASSERT_EQ(doc->GetText().Trim(), System::String::Format(u"Hello world!{0}", ControlChar::Cr()) + u"Hello again!");
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
    builder->Write(System::String(u"Before space.") + ControlChar::SpaceChar + u"After space.");
    
    // Add an NBSP, which is a non-breaking space.
    // Unlike the regular space, this space cannot have an automatic line break at its position.
    builder->Write(System::String(u"Before space.") + ControlChar::NonBreakingSpace() + u"After space.");
    
    // Add a tab character.
    builder->Write(System::String(u"Before tab.") + ControlChar::Tab() + u"After tab.");
    
    // Add a line break.
    builder->Write(System::String(u"Before line break.") + ControlChar::LineBreak() + u"After line break.");
    
    // Add a new line and starts a new paragraph.
    ASSERT_EQ(doc->get_FirstSection()->get_Body()->GetChildNodes(NodeType::Paragraph, true)->get_Count(), 1);
    builder->Write(System::String(u"Before line feed.") + ControlChar::LineFeed() + u"After line feed.");
    ASSERT_EQ(doc->get_FirstSection()->get_Body()->GetChildNodes(NodeType::Paragraph, true)->get_Count(), 2);
    
    // The line feed character has two versions.
    ASSERT_EQ(ControlChar::Lf(), ControlChar::LineFeed());
    
    // Carriage returns and line feeds can be represented together by one character.
    ASSERT_EQ(ControlChar::Cr() + ControlChar::Lf(), ControlChar::CrLf());
    
    // Add a paragraph break, which will start a new paragraph.
    builder->Write(System::String(u"Before paragraph break.") + ControlChar::ParagraphBreak() + u"After paragraph break.");
    ASSERT_EQ(doc->get_FirstSection()->get_Body()->GetChildNodes(NodeType::Paragraph, true)->get_Count(), 3);
    
    // Add a section break. This does not make a new section or paragraph.
    ASSERT_EQ(doc->get_Sections()->get_Count(), 1);
    builder->Write(System::String(u"Before section break.") + ControlChar::SectionBreak() + u"After section break.");
    ASSERT_EQ(doc->get_Sections()->get_Count(), 1);
    
    // Add a page break.
    builder->Write(System::String(u"Before page break.") + ControlChar::PageBreak() + u"After page break.");
    
    // A page break is the same value as a section break.
    ASSERT_EQ(ControlChar::SectionBreak(), ControlChar::PageBreak());
    
    // Insert a new section, and then set its column count to two.
    doc->AppendChild<System::SharedPtr<Aspose::Words::Section>>(System::MakeObject<Aspose::Words::Section>(doc));
    builder->MoveToSection(1);
    builder->get_CurrentSection()->get_PageSetup()->get_TextColumns()->SetCount(2);
    
    // We can use a control character to mark the point where text moves to the next column.
    builder->Write(System::String(u"Text at end of column 1.") + ControlChar::ColumnBreak() + u"Text at beginning of column 2.");
    
    doc->Save(get_ArtifactsDir() + u"ControlChar.InsertControlChars.docx");
    
    // There are char and string counterparts for most characters.
    ASPOSE_ASSERT_EQ(ControlChar::CellChar, System::Convert::ToChar(ControlChar::Cell()));
    ASPOSE_ASSERT_EQ(ControlChar::NonBreakingSpaceChar, System::Convert::ToChar(ControlChar::NonBreakingSpace()));
    ASPOSE_ASSERT_EQ(ControlChar::TabChar, System::Convert::ToChar(ControlChar::Tab()));
    ASPOSE_ASSERT_EQ(ControlChar::LineBreakChar, System::Convert::ToChar(ControlChar::LineBreak()));
    ASPOSE_ASSERT_EQ(ControlChar::LineFeedChar, System::Convert::ToChar(ControlChar::LineFeed()));
    ASPOSE_ASSERT_EQ(ControlChar::ParagraphBreakChar, System::Convert::ToChar(ControlChar::ParagraphBreak()));
    ASPOSE_ASSERT_EQ(ControlChar::SectionBreakChar, System::Convert::ToChar(ControlChar::SectionBreak()));
    ASPOSE_ASSERT_EQ(ControlChar::SectionBreakChar, System::Convert::ToChar(ControlChar::PageBreak()));
    ASPOSE_ASSERT_EQ(ControlChar::ColumnBreakChar, System::Convert::ToChar(ControlChar::ColumnBreak()));
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
