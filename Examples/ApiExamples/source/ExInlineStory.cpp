// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExInlineStory.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/date_time.h>
#include <system/collections/list.h>
#include <system/collections/ienumerable.h>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Sections/StoryType.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Numbering/NumberStyle.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteSeparatorType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteSeparatorCollection.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteSeparator.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteOptions.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteNumberingRule.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Footnotes/EndnoteOptions.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

#include "TestUtil.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1273299660u, ::Aspose::Words::ApiExamples::ExInlineStory, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExInlineStory : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExInlineStory> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExInlineStory>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExInlineStory> ExInlineStory::s_instance;

} // namespace gtest_test

void ExInlineStory::PositionFootnote(Aspose::Words::Notes::FootnotePosition footnotePosition)
{
    //ExStart
    //ExFor:Document.FootnoteOptions
    //ExFor:FootnoteOptions
    //ExFor:FootnoteOptions.Position
    //ExFor:FootnotePosition
    //ExSummary:Shows how to select a different place where the document collects and displays its footnotes.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // A footnote is a way to attach a reference or a side comment to text
    // that does not interfere with the main body text's flow.  
    // Inserting a footnote adds a small superscript reference symbol
    // at the main body text where we insert the footnote.
    // Each footnote also creates an entry at the bottom of the page, consisting of a symbol
    // that matches the reference symbol in the main body text.
    // The reference text that we pass to the document builder's "InsertFootnote" method.
    builder->Write(u"Hello world!");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote contents.");
    
    // We can use the "Position" property to determine where the document will place all its footnotes.
    // If we set the value of the "Position" property to "FootnotePosition.BottomOfPage",
    // every footnote will show up at the bottom of the page that contains its reference mark. This is the default value.
    // If we set the value of the "Position" property to "FootnotePosition.BeneathText",
    // every footnote will show up at the end of the page's text that contains its reference mark.
    doc->get_FootnoteOptions()->set_Position(footnotePosition);
    
    doc->Save(get_ArtifactsDir() + u"InlineStory.PositionFootnote.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"InlineStory.PositionFootnote.docx");
    
    ASSERT_EQ(footnotePosition, doc->get_FootnoteOptions()->get_Position());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote contents.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
}

namespace gtest_test
{

using ExInlineStory_PositionFootnote_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExInlineStory::PositionFootnote)>::type;

struct ExInlineStory_PositionFootnote : public ExInlineStory, public Aspose::Words::ApiExamples::ExInlineStory, public ::testing::WithParamInterface<ExInlineStory_PositionFootnote_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Notes::FootnotePosition::BeneathText),
            std::make_tuple(Aspose::Words::Notes::FootnotePosition::BottomOfPage),
        };
    }
};

TEST_P(ExInlineStory_PositionFootnote, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PositionFootnote(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExInlineStory_PositionFootnote, ::testing::ValuesIn(ExInlineStory_PositionFootnote::TestCases()));

} // namespace gtest_test

void ExInlineStory::PositionEndnote(Aspose::Words::Notes::EndnotePosition endnotePosition)
{
    //ExStart
    //ExFor:Document.EndnoteOptions
    //ExFor:EndnoteOptions
    //ExFor:EndnoteOptions.Position
    //ExFor:EndnotePosition
    //ExSummary:Shows how to select a different place where the document collects and displays its endnotes.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // An endnote is a way to attach a reference or a side comment to text
    // that does not interfere with the main body text's flow. 
    // Inserting an endnote adds a small superscript reference symbol
    // at the main body text where we insert the endnote.
    // Each endnote also creates an entry at the end of the document, consisting of a symbol
    // that matches the reference symbol in the main body text.
    // The reference text that we pass to the document builder's "InsertEndnote" method.
    builder->Write(u"Hello world!");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote contents.");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"This is the second section.");
    
    // We can use the "Position" property to determine where the document will place all its endnotes.
    // If we set the value of the "Position" property to "EndnotePosition.EndOfDocument",
    // every footnote will show up in a collection at the end of the document. This is the default value.
    // If we set the value of the "Position" property to "EndnotePosition.EndOfSection",
    // every footnote will show up in a collection at the end of the section whose text contains the endnote's reference mark.
    doc->get_EndnoteOptions()->set_Position(endnotePosition);
    
    doc->Save(get_ArtifactsDir() + u"InlineStory.PositionEndnote.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"InlineStory.PositionEndnote.docx");
    
    ASSERT_EQ(endnotePosition, doc->get_EndnoteOptions()->get_Position());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, true, System::String::Empty, u"Endnote contents.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
}

namespace gtest_test
{

using ExInlineStory_PositionEndnote_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExInlineStory::PositionEndnote)>::type;

struct ExInlineStory_PositionEndnote : public ExInlineStory, public Aspose::Words::ApiExamples::ExInlineStory, public ::testing::WithParamInterface<ExInlineStory_PositionEndnote_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Notes::EndnotePosition::EndOfDocument),
            std::make_tuple(Aspose::Words::Notes::EndnotePosition::EndOfSection),
        };
    }
};

TEST_P(ExInlineStory_PositionEndnote, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PositionEndnote(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExInlineStory_PositionEndnote, ::testing::ValuesIn(ExInlineStory_PositionEndnote::TestCases()));

} // namespace gtest_test

void ExInlineStory::RefMarkNumberStyle()
{
    //ExStart
    //ExFor:Document.EndnoteOptions
    //ExFor:EndnoteOptions
    //ExFor:EndnoteOptions.NumberStyle
    //ExFor:Document.FootnoteOptions
    //ExFor:FootnoteOptions
    //ExFor:FootnoteOptions.NumberStyle
    //ExSummary:Shows how to change the number style of footnote/endnote reference marks.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Footnotes and endnotes are a way to attach a reference or a side comment to text
    // that does not interfere with the main body text's flow. 
    // Inserting a footnote/endnote adds a small superscript reference symbol
    // at the main body text where we insert the footnote/endnote.
    // Each footnote/endnote also creates an entry, which consists of a symbol that matches the reference
    // symbol in the main body text. The reference text that we pass to the document builder's "InsertEndnote" method.
    // Footnote entries, by default, show up at the bottom of each page that contains
    // their reference symbols, and endnotes show up at the end of the document.
    builder->Write(u"Text 1. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote 1.");
    builder->Write(u"Text 2. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote 2.");
    builder->Write(u"Text 3. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote 3.", u"Custom footnote reference mark");
    
    builder->InsertParagraph();
    
    builder->Write(u"Text 1. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote 1.");
    builder->Write(u"Text 2. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote 2.");
    builder->Write(u"Text 3. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote 3.", u"Custom endnote reference mark");
    
    // By default, the reference symbol for each footnote and endnote is its index
    // among all the document's footnotes/endnotes. Each document maintains separate counts
    // for footnotes and for endnotes. By default, footnotes display their numbers using Arabic numerals,
    // and endnotes display their numbers in lowercase Roman numerals.
    ASSERT_EQ(Aspose::Words::NumberStyle::Arabic, doc->get_FootnoteOptions()->get_NumberStyle());
    ASSERT_EQ(Aspose::Words::NumberStyle::LowercaseRoman, doc->get_EndnoteOptions()->get_NumberStyle());
    
    // We can use the "NumberStyle" property to apply custom numbering styles to footnotes and endnotes.
    // This will not affect footnotes/endnotes with custom reference marks.
    doc->get_FootnoteOptions()->set_NumberStyle(Aspose::Words::NumberStyle::UppercaseRoman);
    doc->get_EndnoteOptions()->set_NumberStyle(Aspose::Words::NumberStyle::UppercaseLetter);
    
    doc->Save(get_ArtifactsDir() + u"InlineStory.RefMarkNumberStyle.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"InlineStory.RefMarkNumberStyle.docx");
    
    ASSERT_EQ(Aspose::Words::NumberStyle::UppercaseRoman, doc->get_FootnoteOptions()->get_NumberStyle());
    ASSERT_EQ(Aspose::Words::NumberStyle::UppercaseLetter, doc->get_EndnoteOptions()->get_NumberStyle());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote 1.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote 2.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 1, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, false, u"Custom footnote reference mark", u"Custom footnote reference mark Footnote 3.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 2, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, true, System::String::Empty, u"Endnote 1.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 3, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, true, System::String::Empty, u"Endnote 2.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 4, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, false, u"Custom endnote reference mark", u"Custom endnote reference mark Endnote 3.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 5, true)));
}

namespace gtest_test
{

TEST_F(ExInlineStory, RefMarkNumberStyle)
{
    s_instance->RefMarkNumberStyle();
}

} // namespace gtest_test

void ExInlineStory::NumberingRule()
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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Footnotes and endnotes are a way to attach a reference or a side comment to text
    // that does not interfere with the main body text's flow. 
    // Inserting a footnote/endnote adds a small superscript reference symbol
    // at the main body text where we insert the footnote/endnote.
    // Each footnote/endnote also creates an entry, which consists of a symbol that matches the reference
    // symbol in the main body text. The reference text that we pass to the document builder's "InsertEndnote" method.
    // Footnote entries, by default, show up at the bottom of each page that contains
    // their reference symbols, and endnotes show up at the end of the document.
    builder->Write(u"Text 1. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote 1.");
    builder->Write(u"Text 2. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote 2.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Write(u"Text 3. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote 3.");
    builder->Write(u"Text 4. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote 4.");
    
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    builder->Write(u"Text 1. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote 1.");
    builder->Write(u"Text 2. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote 2.");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Text 3. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote 3.");
    builder->Write(u"Text 4. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote 4.");
    
    // By default, the reference symbol for each footnote and endnote is its index
    // among all the document's footnotes/endnotes. Each document maintains separate counts
    // for footnotes and endnotes and does not restart these counts at any point.
    ASSERT_EQ(doc->get_FootnoteOptions()->get_RestartRule(), Aspose::Words::Notes::FootnoteNumberingRule::Default);
    ASSERT_EQ(Aspose::Words::Notes::FootnoteNumberingRule::Default, Aspose::Words::Notes::FootnoteNumberingRule::Continuous);
    
    // We can use the "RestartRule" property to get the document to restart
    // the footnote/endnote counts at a new page or section.
    doc->get_FootnoteOptions()->set_RestartRule(Aspose::Words::Notes::FootnoteNumberingRule::RestartPage);
    doc->get_EndnoteOptions()->set_RestartRule(Aspose::Words::Notes::FootnoteNumberingRule::RestartSection);
    
    doc->Save(get_ArtifactsDir() + u"InlineStory.NumberingRule.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"InlineStory.NumberingRule.docx");
    
    ASSERT_EQ(Aspose::Words::Notes::FootnoteNumberingRule::RestartPage, doc->get_FootnoteOptions()->get_RestartRule());
    ASSERT_EQ(Aspose::Words::Notes::FootnoteNumberingRule::RestartSection, doc->get_EndnoteOptions()->get_RestartRule());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote 1.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote 2.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 1, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote 3.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 2, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote 4.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 3, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, true, System::String::Empty, u"Endnote 1.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 4, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, true, System::String::Empty, u"Endnote 2.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 5, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, true, System::String::Empty, u"Endnote 3.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 6, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, true, System::String::Empty, u"Endnote 4.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 7, true)));
}

namespace gtest_test
{

TEST_F(ExInlineStory, NumberingRule)
{
    s_instance->NumberingRule();
}

} // namespace gtest_test

void ExInlineStory::StartNumber()
{
    //ExStart
    //ExFor:Document.EndnoteOptions
    //ExFor:EndnoteOptions
    //ExFor:EndnoteOptions.StartNumber
    //ExFor:Document.FootnoteOptions
    //ExFor:FootnoteOptions
    //ExFor:FootnoteOptions.StartNumber
    //ExSummary:Shows how to set a number at which the document begins the footnote/endnote count.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
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
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote 1.");
    builder->Write(u"Text 2. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote 2.");
    builder->Write(u"Text 3. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote 3.");
    
    builder->InsertParagraph();
    
    builder->Write(u"Text 1. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote 1.");
    builder->Write(u"Text 2. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote 2.");
    builder->Write(u"Text 3. ");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote 3.");
    
    // By default, the reference symbol for each footnote and endnote is its index
    // among all the document's footnotes/endnotes. Each document maintains separate counts
    // for footnotes and for endnotes, which both begin at 1.
    ASSERT_EQ(1, doc->get_FootnoteOptions()->get_StartNumber());
    ASSERT_EQ(1, doc->get_EndnoteOptions()->get_StartNumber());
    
    // We can use the "StartNumber" property to get the document to
    // begin a footnote or endnote count at a different number.
    doc->get_EndnoteOptions()->set_NumberStyle(Aspose::Words::NumberStyle::Arabic);
    doc->get_EndnoteOptions()->set_StartNumber(50);
    
    doc->Save(get_ArtifactsDir() + u"InlineStory.StartNumber.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"InlineStory.StartNumber.docx");
    
    ASSERT_EQ(1, doc->get_FootnoteOptions()->get_StartNumber());
    ASSERT_EQ(50, doc->get_EndnoteOptions()->get_StartNumber());
    ASSERT_EQ(Aspose::Words::NumberStyle::Arabic, doc->get_FootnoteOptions()->get_NumberStyle());
    ASSERT_EQ(Aspose::Words::NumberStyle::Arabic, doc->get_EndnoteOptions()->get_NumberStyle());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote 1.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote 2.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 1, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote 3.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 2, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, true, System::String::Empty, u"Endnote 1.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 3, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, true, System::String::Empty, u"Endnote 2.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 4, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, true, System::String::Empty, u"Endnote 3.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 5, true)));
}

namespace gtest_test
{

TEST_F(ExInlineStory, StartNumber)
{
    s_instance->StartNumber();
}

} // namespace gtest_test

void ExInlineStory::AddFootnote()
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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Add text, and reference it with a footnote. This footnote will place a small superscript reference
    // mark after the text that it references and create an entry below the main body text at the bottom of the page.
    // This entry will contain the footnote's reference mark and the reference text,
    // which we will pass to the document builder's "InsertFootnote" method.
    builder->Write(u"Main body text.");
    System::SharedPtr<Aspose::Words::Notes::Footnote> footnote = builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote text.");
    
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
    footnote = builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote text.");
    
    // We can set a custom reference mark which the footnote will use instead of its index number.
    footnote->set_ReferenceMark(u"RefMark");
    
    ASSERT_FALSE(footnote->get_IsAuto());
    
    // A bookmark with the "IsAuto" flag set to true will still show its real index
    // even if previous bookmarks display custom reference marks, so this bookmark's reference mark will be a "3".
    builder->Write(u" More main body text.");
    footnote = builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote text.");
    
    ASSERT_TRUE(footnote->get_IsAuto());
    
    doc->Save(get_ArtifactsDir() + u"InlineStory.AddFootnote.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"InlineStory.AddFootnote.docx");
    
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote text. More text added by a DocumentBuilder.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, false, u"RefMark", u"Footnote text.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 1, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote text.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 2, true)));
}

namespace gtest_test
{

TEST_F(ExInlineStory, AddFootnote)
{
    s_instance->AddFootnote();
}

} // namespace gtest_test

void ExInlineStory::FootnoteEndnote()
{
    //ExStart
    //ExFor:Footnote.FootnoteType
    //ExSummary:Shows the difference between footnotes and endnotes.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are two ways of attaching numbered references to the text. Both these references will add a
    // small superscript reference mark at the location that we insert them.
    // The reference mark, by default, is the index number of the reference among all the references in the document.
    // Each reference will also create an entry, which will have the same reference mark as in the body text
    // and reference text, which we will pass to the document builder's "InsertFootnote" method.
    // 1 -  A footnote, whose entry will appear on the same page as the text that it references:
    builder->Write(u"Footnote referenced main body text.");
    System::SharedPtr<Aspose::Words::Notes::Footnote> footnote = builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote text, will appear at the bottom of the page that contains the referenced text.");
    
    // 2 -  An endnote, whose entry will appear at the end of the document:
    builder->Write(u"Endnote referenced main body text.");
    System::SharedPtr<Aspose::Words::Notes::Footnote> endnote = builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Endnote text, will appear at the very end of the document.");
    
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    
    ASSERT_EQ(Aspose::Words::Notes::FootnoteType::Footnote, footnote->get_FootnoteType());
    ASSERT_EQ(Aspose::Words::Notes::FootnoteType::Endnote, endnote->get_FootnoteType());
    
    doc->Save(get_ArtifactsDir() + u"InlineStory.FootnoteEndnote.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"InlineStory.FootnoteEndnote.docx");
    
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, u"Footnote text, will appear at the bottom of the page that contains the referenced text.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Endnote, true, System::String::Empty, u"Endnote text, will appear at the very end of the document.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 1, true)));
}

namespace gtest_test
{

TEST_F(ExInlineStory, FootnoteEndnote)
{
    s_instance->FootnoteEndnote();
}

} // namespace gtest_test

void ExInlineStory::AddComment()
{
    //ExStart
    //ExFor:Comment
    //ExFor:InlineStory
    //ExFor:InlineStory.Paragraphs
    //ExFor:InlineStory.FirstParagraph
    //ExFor:Comment.#ctor(DocumentBase, String, String, DateTime)
    //ExSummary:Shows how to add a comment to a paragraph.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Write(u"Hello world!");
    
    auto comment = System::MakeObject<Aspose::Words::Comment>(doc, u"John Doe", u"JD", System::DateTime::get_Today());
    builder->get_CurrentParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Comment>>(comment);
    builder->MoveTo(comment->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(System::MakeObject<Aspose::Words::Paragraph>(doc)));
    builder->Write(u"Comment text.");
    
    ASSERT_EQ(System::DateTime::get_Today(), comment->get_DateTime());
    
    // In Microsoft Word, we can right-click this comment in the document body to edit it, or reply to it. 
    doc->Save(get_ArtifactsDir() + u"InlineStory.AddComment.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"InlineStory.AddComment.docx");
    comment = System::ExplicitCast<Aspose::Words::Comment>(doc->GetChild(Aspose::Words::NodeType::Comment, 0, true));
    
    ASSERT_EQ(u"Comment text.\r", comment->GetText());
    ASSERT_EQ(u"John Doe", comment->get_Author());
    ASSERT_EQ(u"JD", comment->get_Initial());
    ASSERT_EQ(System::DateTime::get_Today(), comment->get_DateTime());
}

namespace gtest_test
{

TEST_F(ExInlineStory, AddComment)
{
    s_instance->AddComment();
}

} // namespace gtest_test

void ExInlineStory::InlineStoryRevisions()
{
    //ExStart
    //ExFor:InlineStory.IsDeleteRevision
    //ExFor:InlineStory.IsInsertRevision
    //ExFor:InlineStory.IsMoveFromRevision
    //ExFor:InlineStory.IsMoveToRevision
    //ExSummary:Shows how to view revision-related properties of InlineStory nodes.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Revision footnotes.docx");
    
    // When we edit the document while the "Track Changes" option, found in via Review -> Tracking,
    // is turned on in Microsoft Word, the changes we apply count as revisions.
    // When editing a document using Aspose.Words, we can begin tracking revisions by
    // invoking the document's "StartTrackRevisions" method and stop tracking by using the "StopTrackRevisions" method.
    // We can either accept revisions to assimilate them into the document
    // or reject them to undo and discard the proposed change.
    ASSERT_TRUE(doc->get_HasRevisions());
    
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Notes::Footnote>>> footnotes = doc->GetChildNodes(Aspose::Words::NodeType::Footnote, true)->LINQ_Cast<System::SharedPtr<Aspose::Words::Notes::Footnote> >()->LINQ_ToList();
    
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

namespace gtest_test
{

TEST_F(ExInlineStory, InlineStoryRevisions)
{
    s_instance->InlineStoryRevisions();
}

} // namespace gtest_test

void ExInlineStory::InsertInlineStoryNodes()
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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    System::SharedPtr<Aspose::Words::Notes::Footnote> footnote = builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, nullptr);
    
    // Table nodes have an "EnsureMinimum()" method that makes sure the table has at least one cell.
    auto table = System::MakeObject<Aspose::Words::Tables::Table>(doc);
    table->EnsureMinimum();
    
    // We can place a table inside a footnote, which will make it appear at the referencing page's footer.
    ASSERT_EQ(0, footnote->get_Tables()->get_Count());
    footnote->AppendChild<System::SharedPtr<Aspose::Words::Tables::Table>>(table);
    ASSERT_EQ(1, footnote->get_Tables()->get_Count());
    ASSERT_EQ(Aspose::Words::NodeType::Table, footnote->get_LastChild()->get_NodeType());
    
    // An InlineStory has an "EnsureMinimum()" method as well, but in this case,
    // it makes sure the last child of the node is a paragraph,
    // for us to be able to click and write text easily in Microsoft Word.
    footnote->EnsureMinimum();
    ASSERT_EQ(Aspose::Words::NodeType::Paragraph, footnote->get_LastChild()->get_NodeType());
    
    // Edit the appearance of the anchor, which is the small superscript number
    // in the main text that points to the footnote.
    footnote->get_Font()->set_Name(u"Arial");
    footnote->get_Font()->set_Color(System::Drawing::Color::get_Green());
    
    // All inline story nodes have their respective story types.
    ASSERT_EQ(Aspose::Words::StoryType::Footnotes, footnote->get_StoryType());
    
    // A comment is another type of inline story.
    auto comment = System::ExplicitCast<Aspose::Words::Comment>(builder->get_CurrentParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Comment>>(System::MakeObject<Aspose::Words::Comment>(doc, u"John Doe", u"J. D.", System::DateTime::get_Now())));
    
    // The parent paragraph of an inline story node will be the one from the main document body.
    ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_FirstParagraph(), comment->get_ParentParagraph());
    
    // However, the last paragraph is the one from the comment text contents,
    // which will be outside the main document body in a speech bubble.
    // A comment will not have any child nodes by default,
    // so we can apply the EnsureMinimum() method to place a paragraph here as well.
    ASSERT_TRUE(System::TestTools::IsNull(comment->get_LastParagraph()));
    comment->EnsureMinimum();
    ASSERT_EQ(Aspose::Words::NodeType::Paragraph, comment->get_LastChild()->get_NodeType());
    
    // Once we have a paragraph, we can move the builder to do it and write our comment.
    builder->MoveTo(comment->get_LastParagraph());
    builder->Write(u"My comment.");
    
    ASSERT_EQ(Aspose::Words::StoryType::Comments, comment->get_StoryType());
    
    doc->Save(get_ArtifactsDir() + u"InlineStory.InsertInlineStoryNodes.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"InlineStory.InsertInlineStoryNodes.docx");
    
    footnote = System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType::Footnote, true, System::String::Empty, System::String::Empty, System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
    ASSERT_EQ(u"Arial", footnote->get_Font()->get_Name());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), footnote->get_Font()->get_Color().ToArgb());
    
    comment = System::ExplicitCast<Aspose::Words::Comment>(doc->GetChild(Aspose::Words::NodeType::Comment, 0, true));
    
    ASSERT_EQ(u"My comment.", comment->ToString(Aspose::Words::SaveFormat::Text).Trim());
}

namespace gtest_test
{

TEST_F(ExInlineStory, InsertInlineStoryNodes)
{
    s_instance->InsertInlineStoryNodes();
}

} // namespace gtest_test

void ExInlineStory::DeleteShapes()
{
    //ExStart
    //ExFor:Story
    //ExFor:Story.DeleteShapes
    //ExFor:Story.StoryType
    //ExFor:StoryType
    //ExSummary:Shows how to remove all shapes from a node.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Use a DocumentBuilder to insert a shape. This is an inline shape,
    // which has a parent Paragraph, which is a child node of the first section's Body.
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::Cube, 100.0, 100.0);
    
    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    
    // We can delete all shapes from the child paragraphs of this Body.
    ASSERT_EQ(Aspose::Words::StoryType::MainText, doc->get_FirstSection()->get_Body()->get_StoryType());
    doc->get_FirstSection()->get_Body()->DeleteShapes();
    
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExInlineStory, DeleteShapes)
{
    s_instance->DeleteShapes();
}

} // namespace gtest_test

void ExInlineStory::UpdateActualReferenceMarks()
{
    //ExStart:UpdateActualReferenceMarks
    //GistId:a775441ecb396eea917a2717cb9e8f8f
    //ExFor:Document.UpdateActualReferenceMarks
    //ExFor:Footnote.ActualReferenceMark
    //ExSummary:Shows how to get actual footnote reference mark.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Footnotes and endnotes.docx");
    
    auto footnote = System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 1, true));
    doc->UpdateFields();
    doc->UpdateActualReferenceMarks();
    
    ASSERT_EQ(u"1", footnote->get_ActualReferenceMark());
    //ExEnd:UpdateActualReferenceMarks
}

namespace gtest_test
{

TEST_F(ExInlineStory, UpdateActualReferenceMarks)
{
    s_instance->UpdateActualReferenceMarks();
}

} // namespace gtest_test

void ExInlineStory::EndnoteSeparator()
{
    //ExStart:EndnoteSeparator
    //GistId:e06aa7a168b57907a5598e823a22bf0a
    //ExFor:DocumentBase.FootnoteSeparators
    //ExFor:FootnoteSeparatorType
    //ExSummary:Shows how to remove endnote separator.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Footnotes and endnotes.docx");
    
    System::SharedPtr<Aspose::Words::Notes::FootnoteSeparator> endnoteSeparator = doc->get_FootnoteSeparators()->idx_get(Aspose::Words::Notes::FootnoteSeparatorType::EndnoteSeparator);
    // Remove endnote separator.
    endnoteSeparator->get_FirstParagraph()->get_FirstChild()->Remove();
    //ExEnd:EndnoteSeparator
    
    doc->Save(get_ArtifactsDir() + u"InlineStory.EndnoteSeparator.docx");
}

namespace gtest_test
{

TEST_F(ExInlineStory, EndnoteSeparator)
{
    s_instance->EndnoteSeparator();
}

} // namespace gtest_test

void ExInlineStory::FootnoteSeparator()
{
    //ExStart:FootnoteSeparator
    //GistId:e06aa7a168b57907a5598e823a22bf0a
    //ExFor:DocumentBase.FootnoteSeparators
    //ExFor:FootnoteSeparator
    //ExFor:FootnoteSeparatorType
    //ExFor:FootnoteSeparatorCollection
    //ExFor:FootnoteSeparatorCollection.Item(FootnoteSeparatorType)
    //ExSummary:Shows how to manage footnote separator format.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Footnotes and endnotes.docx");
    
    System::SharedPtr<Aspose::Words::Notes::FootnoteSeparator> footnoteSeparator = doc->get_FootnoteSeparators()->idx_get(Aspose::Words::Notes::FootnoteSeparatorType::FootnoteSeparator);
    // Align footnote separator.
    footnoteSeparator->get_FirstParagraph()->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    //ExEnd:FootnoteSeparator
    
    doc->Save(get_ArtifactsDir() + u"InlineStory.FootnoteSeparator.docx");
}

namespace gtest_test
{

TEST_F(ExInlineStory, FootnoteSeparator)
{
    s_instance->FootnoteSeparator();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
