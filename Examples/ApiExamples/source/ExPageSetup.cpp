// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExPageSetup.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/enumerator_adapter.h>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/TextOrientation.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Settings/MultiplePagesType.h>
#include <Aspose.Words.Cpp/Model/Sections/TextColumnCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/TextColumn.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionStart.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionLayoutMode.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PaperSize.h>
#include <Aspose.Words.Cpp/Model/Sections/PageVerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/PageBorderDistanceFrom.h>
#include <Aspose.Words.Cpp/Model/Sections/PageBorderAppliesTo.h>
#include <Aspose.Words.Cpp/Model/Sections/Orientation.h>
#include <Aspose.Words.Cpp/Model/Sections/LineNumberRestartMode.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/ChapterPageSeparator.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Numbering/NumberStyle.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnotePosition.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteOptions.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteNumberingRule.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Footnotes/EndnotePosition.h>
#include <Aspose.Words.Cpp/Model/Footnotes/EndnoteOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderType.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>

#include "DocumentHelper.h"


using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Settings;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2651189598u, ::Aspose::Words::ApiExamples::ExPageSetup, ThisTypeBaseTypesInfo);

void ExPageSetup::InsertSectionWithEndnote(System::SharedPtr<Aspose::Words::Document> doc, System::String sectionBodyText, System::String endnoteText)
{
    auto section = System::MakeObject<Aspose::Words::Section>(doc);
    
    doc->AppendChild<System::SharedPtr<Aspose::Words::Section>>(section);
    
    auto body = System::MakeObject<Aspose::Words::Body>(doc);
    section->AppendChild<System::SharedPtr<Aspose::Words::Body>>(body);
    
    ASPOSE_ASSERT_EQ(section, body->get_ParentNode());
    
    auto para = System::MakeObject<Aspose::Words::Paragraph>(doc);
    body->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(para);
    
    ASPOSE_ASSERT_EQ(body, para->get_ParentNode());
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->MoveTo(para);
    builder->Write(sectionBodyText);
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, endnoteText);
}

void ExPageSetup::TestSuppressEndnotes(System::SharedPtr<Aspose::Words::Document> doc)
{
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_Sections()->idx_get(1)->get_PageSetup();
    
    ASSERT_TRUE(pageSetup->get_SuppressEndnotes());
}


namespace gtest_test
{

class ExPageSetup : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExPageSetup> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExPageSetup>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExPageSetup> ExPageSetup::s_instance;

} // namespace gtest_test

void ExPageSetup::ClearFormatting()
{
    //ExStart
    //ExFor:DocumentBuilder.PageSetup
    //ExFor:DocumentBuilder.InsertBreak
    //ExFor:DocumentBuilder.Document
    //ExFor:PageSetup
    //ExFor:PageSetup.Orientation
    //ExFor:PageSetup.VerticalAlignment
    //ExFor:PageSetup.ClearFormatting
    //ExFor:Orientation
    //ExFor:PageVerticalAlignment
    //ExFor:BreakType
    //ExSummary:Shows how to apply and revert page setup settings to sections in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Modify the page setup properties for the builder's current section and add text.
    builder->get_PageSetup()->set_Orientation(Aspose::Words::Orientation::Landscape);
    builder->get_PageSetup()->set_VerticalAlignment(Aspose::Words::PageVerticalAlignment::Center);
    builder->Writeln(u"This is the first section, which landscape oriented with vertically centered text.");
    
    // If we start a new section using a document builder,
    // it will inherit the builder's current page setup properties.
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    
    ASSERT_EQ(Aspose::Words::Orientation::Landscape, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_Orientation());
    ASSERT_EQ(Aspose::Words::PageVerticalAlignment::Center, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_VerticalAlignment());
    
    // We can revert its page setup properties to their default values using the "ClearFormatting" method.
    builder->get_PageSetup()->ClearFormatting();
    
    ASSERT_EQ(Aspose::Words::Orientation::Portrait, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_Orientation());
    ASSERT_EQ(Aspose::Words::PageVerticalAlignment::Top, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_VerticalAlignment());
    
    builder->Writeln(u"This is the second section, which is in default Letter paper size, portrait orientation and top alignment.");
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.ClearFormatting.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.ClearFormatting.docx");
    
    ASSERT_EQ(Aspose::Words::Orientation::Landscape, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_Orientation());
    ASSERT_EQ(Aspose::Words::PageVerticalAlignment::Center, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_VerticalAlignment());
    
    ASSERT_EQ(Aspose::Words::Orientation::Portrait, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_Orientation());
    ASSERT_EQ(Aspose::Words::PageVerticalAlignment::Top, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_VerticalAlignment());
}

namespace gtest_test
{

TEST_F(ExPageSetup, ClearFormatting)
{
    s_instance->ClearFormatting();
}

} // namespace gtest_test

void ExPageSetup::DifferentFirstPageHeaderFooter(bool differentFirstPageHeaderFooter)
{
    //ExStart
    //ExFor:PageSetup.DifferentFirstPageHeaderFooter
    //ExSummary:Shows how to enable or disable primary headers/footers.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are two types of header/footers.
    // 1 -  The "First" header/footer, which appears on the first page of the section.
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderFirst);
    builder->Writeln(u"First page header.");
    
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterFirst);
    builder->Writeln(u"First page footer.");
    
    // 2 -  The "Primary" header/footer, which appears on every page in the section.
    // We can override the primary header/footer by a first and an even page header/footer. 
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Writeln(u"Primary header.");
    
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterPrimary);
    builder->Writeln(u"Primary footer.");
    
    builder->MoveToSection(0);
    builder->Writeln(u"Page 1.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 2.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 3.");
    
    // Each section has a "PageSetup" object that specifies page appearance-related properties
    // such as orientation, size, and borders.
    // Set the "DifferentFirstPageHeaderFooter" property to "true" to apply the first header/footer to the first page.
    // Set the "DifferentFirstPageHeaderFooter" property to "false"
    // to make the first page display the primary header/footer.
    builder->get_PageSetup()->set_DifferentFirstPageHeaderFooter(differentFirstPageHeaderFooter);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.DifferentFirstPageHeaderFooter.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.DifferentFirstPageHeaderFooter.docx");
    
    ASPOSE_ASSERT_EQ(differentFirstPageHeaderFooter, doc->get_FirstSection()->get_PageSetup()->get_DifferentFirstPageHeaderFooter());
}

namespace gtest_test
{

using ExPageSetup_DifferentFirstPageHeaderFooter_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExPageSetup::DifferentFirstPageHeaderFooter)>::type;

struct ExPageSetup_DifferentFirstPageHeaderFooter : public ExPageSetup, public Aspose::Words::ApiExamples::ExPageSetup, public ::testing::WithParamInterface<ExPageSetup_DifferentFirstPageHeaderFooter_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPageSetup_DifferentFirstPageHeaderFooter, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DifferentFirstPageHeaderFooter(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPageSetup_DifferentFirstPageHeaderFooter, ::testing::ValuesIn(ExPageSetup_DifferentFirstPageHeaderFooter::TestCases()));

} // namespace gtest_test

void ExPageSetup::OddAndEvenPagesHeaderFooter(bool oddAndEvenPagesHeaderFooter)
{
    //ExStart
    //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
    //ExSummary:Shows how to enable or disable even page headers/footers.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are two types of header/footers.
    // 1 -  The "Primary" header/footer, which appears on every page in the section.
    // We can override the primary header/footer by a first and an even page header/footer. 
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Writeln(u"Primary header.");
    
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterPrimary);
    builder->Writeln(u"Primary footer.");
    
    // 2 -  The "Even" header/footer, which appears on every even page of this section.
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderEven);
    builder->Writeln(u"Even page header.");
    
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterEven);
    builder->Writeln(u"Even page footer.");
    
    builder->MoveToSection(0);
    builder->Writeln(u"Page 1.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 2.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 3.");
    
    // Each section has a "PageSetup" object that specifies page appearance-related properties
    // such as orientation, size, and borders.
    // Set the "OddAndEvenPagesHeaderFooter" property to "true"
    // to display the even page header/footer on even pages.
    // Set the "OddAndEvenPagesHeaderFooter" property to "false"
    // to display the primary header/footer on even pages.
    builder->get_PageSetup()->set_OddAndEvenPagesHeaderFooter(oddAndEvenPagesHeaderFooter);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.OddAndEvenPagesHeaderFooter.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.OddAndEvenPagesHeaderFooter.docx");
    
    ASPOSE_ASSERT_EQ(oddAndEvenPagesHeaderFooter, doc->get_FirstSection()->get_PageSetup()->get_OddAndEvenPagesHeaderFooter());
}

namespace gtest_test
{

using ExPageSetup_OddAndEvenPagesHeaderFooter_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExPageSetup::OddAndEvenPagesHeaderFooter)>::type;

struct ExPageSetup_OddAndEvenPagesHeaderFooter : public ExPageSetup, public Aspose::Words::ApiExamples::ExPageSetup, public ::testing::WithParamInterface<ExPageSetup_OddAndEvenPagesHeaderFooter_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPageSetup_OddAndEvenPagesHeaderFooter, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->OddAndEvenPagesHeaderFooter(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPageSetup_OddAndEvenPagesHeaderFooter, ::testing::ValuesIn(ExPageSetup_OddAndEvenPagesHeaderFooter::TestCases()));

} // namespace gtest_test

void ExPageSetup::CharactersPerLine()
{
    //ExStart
    //ExFor:PageSetup.CharactersPerLine
    //ExFor:PageSetup.LayoutMode
    //ExFor:SectionLayoutMode
    //ExSummary:Shows how to specify a for the number of characters that each line may have.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Enable pitching, and then use it to set the number of characters per line in this section.
    builder->get_PageSetup()->set_LayoutMode(Aspose::Words::SectionLayoutMode::Grid);
    builder->get_PageSetup()->set_CharactersPerLine(10);
    
    // The number of characters also depends on the size of the font.
    doc->get_Styles()->idx_get(u"Normal")->get_Font()->set_Size(20);
    
    ASSERT_EQ(8, doc->get_FirstSection()->get_PageSetup()->get_CharactersPerLine());
    
    builder->Writeln(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.CharactersPerLine.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.CharactersPerLine.docx");
    
    ASSERT_EQ(Aspose::Words::SectionLayoutMode::Grid, doc->get_FirstSection()->get_PageSetup()->get_LayoutMode());
    ASSERT_EQ(8, doc->get_FirstSection()->get_PageSetup()->get_CharactersPerLine());
}

namespace gtest_test
{

TEST_F(ExPageSetup, CharactersPerLine)
{
    s_instance->CharactersPerLine();
}

} // namespace gtest_test

void ExPageSetup::LinesPerPage()
{
    //ExStart
    //ExFor:PageSetup.LinesPerPage
    //ExFor:PageSetup.LayoutMode
    //ExFor:ParagraphFormat.SnapToGrid
    //ExFor:SectionLayoutMode
    //ExSummary:Shows how to specify a limit for the number of lines that each page may have.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Enable pitching, and then use it to set the number of lines per page in this section.
    // A large enough font size will push some lines down onto the next page to avoid overlapping characters.
    builder->get_PageSetup()->set_LayoutMode(Aspose::Words::SectionLayoutMode::LineGrid);
    builder->get_PageSetup()->set_LinesPerPage(15);
    
    builder->get_ParagraphFormat()->set_SnapToGrid(true);
    
    for (int32_t i = 0; i < 30; i++)
    {
        builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");
    }
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.LinesPerPage.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.LinesPerPage.docx");
    
    ASSERT_EQ(Aspose::Words::SectionLayoutMode::LineGrid, doc->get_FirstSection()->get_PageSetup()->get_LayoutMode());
    ASSERT_EQ(15, doc->get_FirstSection()->get_PageSetup()->get_LinesPerPage());
    
    for (auto&& paragraph : System::IterateOver<Aspose::Words::Paragraph>(doc->get_FirstSection()->get_Body()->get_Paragraphs()))
    {
        ASSERT_TRUE(paragraph->get_ParagraphFormat()->get_SnapToGrid());
    }
}

namespace gtest_test
{

TEST_F(ExPageSetup, LinesPerPage)
{
    s_instance->LinesPerPage();
}

} // namespace gtest_test

void ExPageSetup::SetSectionStart()
{
    //ExStart
    //ExFor:SectionStart
    //ExFor:PageSetup.SectionStart
    //ExFor:Document.Sections
    //ExSummary:Shows how to specify how a new section separates itself from the previous.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"This text is in section 1.");
    
    // Section break types determine how a new section separates itself from the previous section.
    // Below are five types of section breaks.
    // 1 -  Starts the next section on a new page:
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Writeln(u"This text is in section 2.");
    
    ASSERT_EQ(Aspose::Words::SectionStart::NewPage, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_SectionStart());
    
    // 2 -  Starts the next section on the current page:
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakContinuous);
    builder->Writeln(u"This text is in section 3.");
    
    ASSERT_EQ(Aspose::Words::SectionStart::Continuous, doc->get_Sections()->idx_get(2)->get_PageSetup()->get_SectionStart());
    
    // 3 -  Starts the next section on a new even page:
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakEvenPage);
    builder->Writeln(u"This text is in section 4.");
    
    ASSERT_EQ(Aspose::Words::SectionStart::EvenPage, doc->get_Sections()->idx_get(3)->get_PageSetup()->get_SectionStart());
    
    // 4 -  Starts the next section on a new odd page:
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakOddPage);
    builder->Writeln(u"This text is in section 5.");
    
    ASSERT_EQ(Aspose::Words::SectionStart::OddPage, doc->get_Sections()->idx_get(4)->get_PageSetup()->get_SectionStart());
    
    // 5 -  Starts the next section on a new column:
    System::SharedPtr<Aspose::Words::TextColumnCollection> columns = builder->get_PageSetup()->get_TextColumns();
    columns->SetCount(2);
    
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewColumn);
    builder->Writeln(u"This text is in section 6.");
    
    ASSERT_EQ(Aspose::Words::SectionStart::NewColumn, doc->get_Sections()->idx_get(5)->get_PageSetup()->get_SectionStart());
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.SetSectionStart.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.SetSectionStart.docx");
    
    ASSERT_EQ(Aspose::Words::SectionStart::NewPage, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_SectionStart());
    ASSERT_EQ(Aspose::Words::SectionStart::NewPage, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_SectionStart());
    ASSERT_EQ(Aspose::Words::SectionStart::Continuous, doc->get_Sections()->idx_get(2)->get_PageSetup()->get_SectionStart());
    ASSERT_EQ(Aspose::Words::SectionStart::EvenPage, doc->get_Sections()->idx_get(3)->get_PageSetup()->get_SectionStart());
    ASSERT_EQ(Aspose::Words::SectionStart::OddPage, doc->get_Sections()->idx_get(4)->get_PageSetup()->get_SectionStart());
    ASSERT_EQ(Aspose::Words::SectionStart::NewColumn, doc->get_Sections()->idx_get(5)->get_PageSetup()->get_SectionStart());
}

namespace gtest_test
{

TEST_F(ExPageSetup, SetSectionStart)
{
    s_instance->SetSectionStart();
}

} // namespace gtest_test

void ExPageSetup::PageMargins()
{
    //ExStart
    //ExFor:ConvertUtil
    //ExFor:ConvertUtil.InchToPoint
    //ExFor:PaperSize
    //ExFor:PageSetup.PaperSize
    //ExFor:PageSetup.Orientation
    //ExFor:PageSetup.TopMargin
    //ExFor:PageSetup.BottomMargin
    //ExFor:PageSetup.LeftMargin
    //ExFor:PageSetup.RightMargin
    //ExFor:PageSetup.HeaderDistance
    //ExFor:PageSetup.FooterDistance
    //ExSummary:Shows how to adjust paper size, orientation, margins, along with other settings for a section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_PageSetup()->set_PaperSize(Aspose::Words::PaperSize::Legal);
    builder->get_PageSetup()->set_Orientation(Aspose::Words::Orientation::Landscape);
    builder->get_PageSetup()->set_TopMargin(Aspose::Words::ConvertUtil::InchToPoint(1.0));
    builder->get_PageSetup()->set_BottomMargin(Aspose::Words::ConvertUtil::InchToPoint(1.0));
    builder->get_PageSetup()->set_LeftMargin(Aspose::Words::ConvertUtil::InchToPoint(1.5));
    builder->get_PageSetup()->set_RightMargin(Aspose::Words::ConvertUtil::InchToPoint(1.5));
    builder->get_PageSetup()->set_HeaderDistance(Aspose::Words::ConvertUtil::InchToPoint(0.2));
    builder->get_PageSetup()->set_FooterDistance(Aspose::Words::ConvertUtil::InchToPoint(0.2));
    
    builder->Writeln(u"Hello world!");
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.PageMargins.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.PageMargins.docx");
    
    ASSERT_EQ(Aspose::Words::PaperSize::Legal, doc->get_FirstSection()->get_PageSetup()->get_PaperSize());
    ASPOSE_ASSERT_EQ(1008.0, doc->get_FirstSection()->get_PageSetup()->get_PageWidth());
    ASPOSE_ASSERT_EQ(612.0, doc->get_FirstSection()->get_PageSetup()->get_PageHeight());
    ASSERT_EQ(Aspose::Words::Orientation::Landscape, doc->get_FirstSection()->get_PageSetup()->get_Orientation());
    ASPOSE_ASSERT_EQ(72.0, doc->get_FirstSection()->get_PageSetup()->get_TopMargin());
    ASPOSE_ASSERT_EQ(72.0, doc->get_FirstSection()->get_PageSetup()->get_BottomMargin());
    ASPOSE_ASSERT_EQ(108.0, doc->get_FirstSection()->get_PageSetup()->get_LeftMargin());
    ASPOSE_ASSERT_EQ(108.0, doc->get_FirstSection()->get_PageSetup()->get_RightMargin());
    ASPOSE_ASSERT_EQ(14.4, doc->get_FirstSection()->get_PageSetup()->get_HeaderDistance());
    ASPOSE_ASSERT_EQ(14.4, doc->get_FirstSection()->get_PageSetup()->get_FooterDistance());
}

namespace gtest_test
{

TEST_F(ExPageSetup, PageMargins)
{
    s_instance->PageMargins();
}

} // namespace gtest_test

void ExPageSetup::PaperSizes()
{
    //ExStart
    //ExFor:PaperSize
    //ExFor:PageSetup.PaperSize
    //ExSummary:Shows how to set page sizes.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // We can change the current page's size to a pre-defined size
    // by using the "PaperSize" property of this section's PageSetup object.
    builder->get_PageSetup()->set_PaperSize(Aspose::Words::PaperSize::Tabloid);
    
    ASPOSE_ASSERT_EQ(792.0, builder->get_PageSetup()->get_PageWidth());
    ASPOSE_ASSERT_EQ(1224.0, builder->get_PageSetup()->get_PageHeight());
    
    builder->Writeln(System::String::Format(u"This page is {0}x{1}.", builder->get_PageSetup()->get_PageWidth(), builder->get_PageSetup()->get_PageHeight()));
    
    // Each section has its own PageSetup object. When we use a document builder to make a new section,
    // that section's PageSetup object inherits all the previous section's PageSetup object's values.
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakEvenPage);
    
    ASSERT_EQ(Aspose::Words::PaperSize::Tabloid, builder->get_PageSetup()->get_PaperSize());
    
    builder->get_PageSetup()->set_PaperSize(Aspose::Words::PaperSize::A5);
    builder->Writeln(System::String::Format(u"This page is {0}x{1}.", builder->get_PageSetup()->get_PageWidth(), builder->get_PageSetup()->get_PageHeight()));
    
    ASPOSE_ASSERT_EQ(419.55, builder->get_PageSetup()->get_PageWidth());
    ASPOSE_ASSERT_EQ(595.30, builder->get_PageSetup()->get_PageHeight());
    
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakEvenPage);
    
    // Set a custom size for this section's pages.
    builder->get_PageSetup()->set_PageWidth(620);
    builder->get_PageSetup()->set_PageHeight(480);
    
    ASSERT_EQ(Aspose::Words::PaperSize::Custom, builder->get_PageSetup()->get_PaperSize());
    
    builder->Writeln(System::String::Format(u"This page is {0}x{1}.", builder->get_PageSetup()->get_PageWidth(), builder->get_PageSetup()->get_PageHeight()));
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.PaperSizes.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.PaperSizes.docx");
    
    ASSERT_EQ(Aspose::Words::PaperSize::Tabloid, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_PaperSize());
    ASPOSE_ASSERT_EQ(792.0, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_PageWidth());
    ASPOSE_ASSERT_EQ(1224.0, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_PageHeight());
    ASSERT_EQ(Aspose::Words::PaperSize::A5, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_PaperSize());
    ASPOSE_ASSERT_EQ(419.55, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_PageWidth());
    ASPOSE_ASSERT_EQ(595.30, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_PageHeight());
    ASSERT_EQ(Aspose::Words::PaperSize::Custom, doc->get_Sections()->idx_get(2)->get_PageSetup()->get_PaperSize());
    ASPOSE_ASSERT_EQ(620.0, doc->get_Sections()->idx_get(2)->get_PageSetup()->get_PageWidth());
    ASPOSE_ASSERT_EQ(480.0, doc->get_Sections()->idx_get(2)->get_PageSetup()->get_PageHeight());
}

namespace gtest_test
{

TEST_F(ExPageSetup, PaperSizes)
{
    s_instance->PaperSizes();
}

} // namespace gtest_test

void ExPageSetup::ColumnsSameWidth()
{
    //ExStart
    //ExFor:PageSetup.TextColumns
    //ExFor:TextColumnCollection
    //ExFor:TextColumnCollection.Spacing
    //ExFor:TextColumnCollection.SetCount
    //ExFor:TextColumnCollection.Count
    //ExFor:TextColumnCollection.Width
    //ExSummary:Shows how to create multiple evenly spaced columns in a section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::TextColumnCollection> columns = builder->get_PageSetup()->get_TextColumns();
    columns->set_Spacing(100);
    columns->SetCount(2);
    
    builder->Writeln(u"Column 1.");
    builder->InsertBreak(Aspose::Words::BreakType::ColumnBreak);
    builder->Writeln(u"Column 2.");
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.ColumnsSameWidth.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.ColumnsSameWidth.docx");
    
    ASPOSE_ASSERT_EQ(100.0, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_Spacing());
    ASSERT_EQ(2, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_Count());
    ASSERT_NEAR(185.15, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_Width(), 0.01);
}

namespace gtest_test
{

TEST_F(ExPageSetup, ColumnsSameWidth)
{
    s_instance->ColumnsSameWidth();
}

} // namespace gtest_test

void ExPageSetup::CustomColumnWidth()
{
    //ExStart
    //ExFor:TextColumnCollection.EvenlySpaced
    //ExFor:TextColumnCollection.Item
    //ExFor:TextColumn
    //ExFor:TextColumn.Width
    //ExFor:TextColumn.SpaceAfter
    //ExSummary:Shows how to create unevenly spaced columns.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = builder->get_PageSetup();
    
    System::SharedPtr<Aspose::Words::TextColumnCollection> columns = pageSetup->get_TextColumns();
    columns->set_EvenlySpaced(false);
    columns->SetCount(2);
    
    // Determine the amount of room that we have available for arranging columns.
    double contentWidth = pageSetup->get_PageWidth() - pageSetup->get_LeftMargin() - pageSetup->get_RightMargin();
    
    ASSERT_NEAR(470.30, contentWidth, 0.01);
    
    // Set the first column to be narrow.
    System::SharedPtr<Aspose::Words::TextColumn> column = columns->idx_get(0);
    column->set_Width(100);
    column->set_SpaceAfter(20);
    
    // Set the second column to take the rest of the space available within the margins of the page.
    column = columns->idx_get(1);
    column->set_Width(contentWidth - column->get_Width() - column->get_SpaceAfter());
    
    builder->Writeln(u"Narrow column 1.");
    builder->InsertBreak(Aspose::Words::BreakType::ColumnBreak);
    builder->Writeln(u"Wide column 2.");
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.CustomColumnWidth.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.CustomColumnWidth.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_FALSE(pageSetup->get_TextColumns()->get_EvenlySpaced());
    ASSERT_EQ(2, pageSetup->get_TextColumns()->get_Count());
    ASPOSE_ASSERT_EQ(100.0, pageSetup->get_TextColumns()->idx_get(0)->get_Width());
    ASPOSE_ASSERT_EQ(20.0, pageSetup->get_TextColumns()->idx_get(0)->get_SpaceAfter());
    ASPOSE_ASSERT_EQ(470.3, pageSetup->get_TextColumns()->idx_get(1)->get_Width());
    ASPOSE_ASSERT_EQ(0.0, pageSetup->get_TextColumns()->idx_get(1)->get_SpaceAfter());
}

namespace gtest_test
{

TEST_F(ExPageSetup, CustomColumnWidth)
{
    s_instance->CustomColumnWidth();
}

} // namespace gtest_test

void ExPageSetup::VerticalLineBetweenColumns(bool lineBetween)
{
    //ExStart
    //ExFor:TextColumnCollection.LineBetween
    //ExSummary:Shows how to separate columns with a vertical line.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Configure the current section's PageSetup object to divide the text into several columns.
    // Set the "LineBetween" property to "true" to put a dividing line between columns.
    // Set the "LineBetween" property to "false" to leave the space between columns blank.
    System::SharedPtr<Aspose::Words::TextColumnCollection> columns = builder->get_PageSetup()->get_TextColumns();
    columns->set_LineBetween(lineBetween);
    columns->SetCount(3);
    
    builder->Writeln(u"Column 1.");
    builder->InsertBreak(Aspose::Words::BreakType::ColumnBreak);
    builder->Writeln(u"Column 2.");
    builder->InsertBreak(Aspose::Words::BreakType::ColumnBreak);
    builder->Writeln(u"Column 3.");
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.VerticalLineBetweenColumns.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.VerticalLineBetweenColumns.docx");
    
    ASPOSE_ASSERT_EQ(lineBetween, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_LineBetween());
}

namespace gtest_test
{

using ExPageSetup_VerticalLineBetweenColumns_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExPageSetup::VerticalLineBetweenColumns)>::type;

struct ExPageSetup_VerticalLineBetweenColumns : public ExPageSetup, public Aspose::Words::ApiExamples::ExPageSetup, public ::testing::WithParamInterface<ExPageSetup_VerticalLineBetweenColumns_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPageSetup_VerticalLineBetweenColumns, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->VerticalLineBetweenColumns(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPageSetup_VerticalLineBetweenColumns, ::testing::ValuesIn(ExPageSetup_VerticalLineBetweenColumns::TestCases()));

} // namespace gtest_test

void ExPageSetup::LineNumbers()
{
    //ExStart
    //ExFor:PageSetup.LineStartingNumber
    //ExFor:PageSetup.LineNumberDistanceFromText
    //ExFor:PageSetup.LineNumberCountBy
    //ExFor:PageSetup.LineNumberRestartMode
    //ExFor:ParagraphFormat.SuppressLineNumbers
    //ExFor:LineNumberRestartMode
    //ExSummary:Shows how to enable line numbering for a section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // We can use the section's PageSetup object to display numbers to the left of the section's text lines.
    // This is the same behavior as a List object,
    // but it covers the entire section and does not modify the text in any way.
    // Our section will restart the numbering on each new page from 1 and display the number,
    // if it is a multiple of 3, at 50pt to the left of the line.
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = builder->get_PageSetup();
    pageSetup->set_LineStartingNumber(1);
    pageSetup->set_LineNumberCountBy(3);
    pageSetup->set_LineNumberRestartMode(Aspose::Words::LineNumberRestartMode::RestartPage);
    pageSetup->set_LineNumberDistanceFromText(50.0);
    
    for (int32_t i = 1; i <= 25; i++)
    {
        builder->Writeln(System::String::Format(u"Line {0}.", i));
    }
    
    // The line counter will skip any paragraph with the "SuppressLineNumbers" flag set to "true".
    // This paragraph is on the 15th line, which is a multiple of 3, and thus would normally display a line number.
    // The section's line counter will also ignore this line, treat the next line as the 15th,
    // and continue the count from that point onward.
    doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(14)->get_ParagraphFormat()->set_SuppressLineNumbers(true);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.LineNumbers.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.LineNumbers.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_EQ(1, pageSetup->get_LineStartingNumber());
    ASSERT_EQ(3, pageSetup->get_LineNumberCountBy());
    ASSERT_EQ(Aspose::Words::LineNumberRestartMode::RestartPage, pageSetup->get_LineNumberRestartMode());
    ASPOSE_ASSERT_EQ(50.0, pageSetup->get_LineNumberDistanceFromText());
}

namespace gtest_test
{

TEST_F(ExPageSetup, LineNumbers)
{
    s_instance->LineNumbers();
}

} // namespace gtest_test

void ExPageSetup::PageBorderProperties()
{
    //ExStart
    //ExFor:Section.PageSetup
    //ExFor:PageSetup.BorderAlwaysInFront
    //ExFor:PageSetup.BorderDistanceFrom
    //ExFor:PageSetup.BorderAppliesTo
    //ExFor:PageBorderDistanceFrom
    //ExFor:PageBorderAppliesTo
    //ExFor:Border.DistanceFromText
    //ExSummary:Shows how to create a wide blue band border at the top of the first page.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->set_BorderAlwaysInFront(false);
    pageSetup->set_BorderDistanceFrom(Aspose::Words::PageBorderDistanceFrom::PageEdge);
    pageSetup->set_BorderAppliesTo(Aspose::Words::PageBorderAppliesTo::FirstPage);
    
    System::SharedPtr<Aspose::Words::Border> border = pageSetup->get_Borders()->idx_get(Aspose::Words::BorderType::Top);
    border->set_LineStyle(Aspose::Words::LineStyle::Single);
    border->set_LineWidth(30);
    border->set_Color(System::Drawing::Color::get_Blue());
    border->set_DistanceFromText(0);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.PageBorderProperties.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.PageBorderProperties.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_FALSE(pageSetup->get_BorderAlwaysInFront());
    ASSERT_EQ(Aspose::Words::PageBorderDistanceFrom::PageEdge, pageSetup->get_BorderDistanceFrom());
    ASSERT_EQ(Aspose::Words::PageBorderAppliesTo::FirstPage, pageSetup->get_BorderAppliesTo());
    
    border = pageSetup->get_Borders()->idx_get(Aspose::Words::BorderType::Top);
    
    ASSERT_EQ(Aspose::Words::LineStyle::Single, border->get_LineStyle());
    ASPOSE_ASSERT_EQ(30.0, border->get_LineWidth());
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), border->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(0.0, border->get_DistanceFromText());
}

namespace gtest_test
{

TEST_F(ExPageSetup, PageBorderProperties)
{
    s_instance->PageBorderProperties();
}

} // namespace gtest_test

void ExPageSetup::PageBorders()
{
    //ExStart
    //ExFor:PageSetup.Borders
    //ExFor:Border.Shadow
    //ExFor:BorderCollection.LineStyle
    //ExFor:BorderCollection.LineWidth
    //ExFor:BorderCollection.Color
    //ExFor:BorderCollection.DistanceFromText
    //ExFor:BorderCollection.Shadow
    //ExSummary:Shows how to create green wavy page border with a shadow.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    
    pageSetup->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::DoubleWave);
    pageSetup->get_Borders()->set_LineWidth(2);
    pageSetup->get_Borders()->set_Color(System::Drawing::Color::get_Green());
    pageSetup->get_Borders()->set_DistanceFromText(24);
    pageSetup->get_Borders()->set_Shadow(true);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.PageBorders.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.PageBorders.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    for (auto&& border : System::IterateOver(pageSetup->get_Borders()))
    {
        ASSERT_EQ(Aspose::Words::LineStyle::DoubleWave, border->get_LineStyle());
        ASPOSE_ASSERT_EQ(2.0, border->get_LineWidth());
        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), border->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(24.0, border->get_DistanceFromText());
        ASSERT_TRUE(border->get_Shadow());
    }
}

namespace gtest_test
{

TEST_F(ExPageSetup, PageBorders)
{
    s_instance->PageBorders();
}

} // namespace gtest_test

void ExPageSetup::PageNumbering()
{
    //ExStart
    //ExFor:PageSetup.RestartPageNumbering
    //ExFor:PageSetup.PageStartingNumber
    //ExFor:PageSetup.PageNumberStyle
    //ExFor:DocumentBuilder.InsertField(String, String)
    //ExSummary:Shows how to set up page numbering in a section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Section 1, page 1.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Section 1, page 2.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Section 1, page 3.");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Writeln(u"Section 2, page 1.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Section 2, page 2.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Section 2, page 3.");
    
    // Move the document builder to the first section's primary header,
    // which every page in that section will display.
    builder->MoveToSection(0);
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    
    // Insert a PAGE field, which will display the number of the current page.
    builder->Write(u"Page ");
    builder->InsertField(u"PAGE", u"");
    
    // Configure the section to have the page count that PAGE fields display start from 5.
    // Also, configure all PAGE fields to display their page numbers using uppercase Roman numerals.
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->set_RestartPageNumbering(true);
    pageSetup->set_PageStartingNumber(5);
    pageSetup->set_PageNumberStyle(Aspose::Words::NumberStyle::UppercaseRoman);
    
    // Create another primary header for the second section, with another PAGE field.
    builder->MoveToSection(1);
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    builder->Write(u" - ");
    builder->InsertField(u"PAGE", u"");
    builder->Write(u" - ");
    
    // Configure the section to have the page count that PAGE fields display start from 10.
    // Also, configure all PAGE fields to display their page numbers using Arabic numbers.
    pageSetup = doc->get_Sections()->idx_get(1)->get_PageSetup();
    pageSetup->set_PageStartingNumber(10);
    pageSetup->set_RestartPageNumbering(true);
    pageSetup->set_PageNumberStyle(Aspose::Words::NumberStyle::Arabic);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.PageNumbering.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.PageNumbering.docx");
    pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    
    ASSERT_TRUE(pageSetup->get_RestartPageNumbering());
    ASSERT_EQ(5, pageSetup->get_PageStartingNumber());
    ASSERT_EQ(Aspose::Words::NumberStyle::UppercaseRoman, pageSetup->get_PageNumberStyle());
    
    pageSetup = doc->get_Sections()->idx_get(1)->get_PageSetup();
    
    ASSERT_TRUE(pageSetup->get_RestartPageNumbering());
    ASSERT_EQ(10, pageSetup->get_PageStartingNumber());
    ASSERT_EQ(Aspose::Words::NumberStyle::Arabic, pageSetup->get_PageNumberStyle());
}

namespace gtest_test
{

TEST_F(ExPageSetup, PageNumbering)
{
    s_instance->PageNumbering();
}

} // namespace gtest_test

void ExPageSetup::FootnoteOptions()
{
    //ExStart
    //ExFor:PageSetup.EndnoteOptions
    //ExFor:PageSetup.FootnoteOptions
    //ExSummary:Shows how to configure options affecting footnotes/endnotes in a section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Hello world!");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Footnote reference text.");
    
    // Configure all footnotes in the first section to restart the numbering from 1
    // at each new page and display themselves directly beneath the text on every page.
    System::SharedPtr<Aspose::Words::Notes::FootnoteOptions> footnoteOptions = doc->get_Sections()->idx_get(0)->get_PageSetup()->get_FootnoteOptions();
    footnoteOptions->set_Position(Aspose::Words::Notes::FootnotePosition::BeneathText);
    footnoteOptions->set_RestartRule(Aspose::Words::Notes::FootnoteNumberingRule::RestartPage);
    footnoteOptions->set_StartNumber(1);
    
    builder->Write(u" Hello again.");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Endnote reference text.");
    
    // Configure all endnotes in the first section to maintain a continuous count throughout the section,
    // starting from 1. Also, set them all to appear collected at the end of the document.
    System::SharedPtr<Aspose::Words::Notes::EndnoteOptions> endnoteOptions = doc->get_Sections()->idx_get(0)->get_PageSetup()->get_EndnoteOptions();
    endnoteOptions->set_Position(Aspose::Words::Notes::EndnotePosition::EndOfDocument);
    endnoteOptions->set_RestartRule(Aspose::Words::Notes::FootnoteNumberingRule::Continuous);
    endnoteOptions->set_StartNumber(1);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.FootnoteOptions.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.FootnoteOptions.docx");
    footnoteOptions = doc->get_FirstSection()->get_PageSetup()->get_FootnoteOptions();
    
    ASSERT_EQ(Aspose::Words::Notes::FootnotePosition::BeneathText, footnoteOptions->get_Position());
    ASSERT_EQ(Aspose::Words::Notes::FootnoteNumberingRule::RestartPage, footnoteOptions->get_RestartRule());
    ASSERT_EQ(1, footnoteOptions->get_StartNumber());
    
    endnoteOptions = doc->get_FirstSection()->get_PageSetup()->get_EndnoteOptions();
    
    ASSERT_EQ(Aspose::Words::Notes::EndnotePosition::EndOfDocument, endnoteOptions->get_Position());
    ASSERT_EQ(Aspose::Words::Notes::FootnoteNumberingRule::Continuous, endnoteOptions->get_RestartRule());
    ASSERT_EQ(1, endnoteOptions->get_StartNumber());
}

namespace gtest_test
{

TEST_F(ExPageSetup, FootnoteOptions)
{
    s_instance->FootnoteOptions();
}

} // namespace gtest_test

void ExPageSetup::Bidi(bool reverseColumns)
{
    //ExStart
    //ExFor:PageSetup.Bidi
    //ExSummary:Shows how to set the order of text columns in a section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->get_TextColumns()->SetCount(3);
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Write(u"Column 1.");
    builder->InsertBreak(Aspose::Words::BreakType::ColumnBreak);
    builder->Write(u"Column 2.");
    builder->InsertBreak(Aspose::Words::BreakType::ColumnBreak);
    builder->Write(u"Column 3.");
    
    // Set the "Bidi" property to "true" to arrange the columns starting from the page's right side.
    // The order of the columns will match the direction of the right-to-left text.
    // Set the "Bidi" property to "false" to arrange the columns starting from the page's left side.
    // The order of the columns will match the direction of the left-to-right text.
    pageSetup->set_Bidi(reverseColumns);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.Bidi.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.Bidi.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_EQ(3, pageSetup->get_TextColumns()->get_Count());
    ASPOSE_ASSERT_EQ(reverseColumns, pageSetup->get_Bidi());
}

namespace gtest_test
{

using ExPageSetup_Bidi_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExPageSetup::Bidi)>::type;

struct ExPageSetup_Bidi : public ExPageSetup, public Aspose::Words::ApiExamples::ExPageSetup, public ::testing::WithParamInterface<ExPageSetup_Bidi_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPageSetup_Bidi, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Bidi(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPageSetup_Bidi, ::testing::ValuesIn(ExPageSetup_Bidi::TestCases()));

} // namespace gtest_test

void ExPageSetup::PageBorder()
{
    //ExStart
    //ExFor:PageSetup.BorderSurroundsFooter
    //ExFor:PageSetup.BorderSurroundsHeader
    //ExSummary:Shows how to apply a border to the page and header/footer.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world! This is the main body text.");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Write(u"This is the header.");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterPrimary);
    builder->Write(u"This is the footer.");
    builder->MoveToDocumentEnd();
    
    // Insert a blue double-line border.
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::Double);
    pageSetup->get_Borders()->set_Color(System::Drawing::Color::get_Blue());
    
    // A section's PageSetup object has "BorderSurroundsHeader" and "BorderSurroundsFooter" flags that determine
    // whether a page border surrounds the main body text, also includes the header or footer, respectively.
    // Set the "BorderSurroundsHeader" flag to "true" to surround the header with our border,
    // and then set the "BorderSurroundsFooter" flag to leave the footer outside of the border.
    pageSetup->set_BorderSurroundsHeader(true);
    pageSetup->set_BorderSurroundsFooter(false);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.PageBorder.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.PageBorder.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_TRUE(pageSetup->get_BorderSurroundsHeader());
    ASSERT_FALSE(pageSetup->get_BorderSurroundsFooter());
}

namespace gtest_test
{

TEST_F(ExPageSetup, PageBorder)
{
    s_instance->PageBorder();
}

} // namespace gtest_test

void ExPageSetup::Gutter()
{
    //ExStart
    //ExFor:PageSetup.Gutter
    //ExFor:PageSetup.RtlGutter
    //ExFor:PageSetup.MultiplePages
    //ExSummary:Shows how to set gutter margins.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Insert text that spans several pages.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    for (int32_t i = 0; i < 6; i++)
    {
        builder->Write(System::String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") + u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    }
    
    // A gutter adds whitespaces to either the left or right page margin,
    // which makes up for the center folding of pages in a book encroaching on the page's layout.
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    
    // Determine how much space our pages have for text within the margins and then add an amount to pad a margin. 
    ASSERT_NEAR(470.30, pageSetup->get_PageWidth() - pageSetup->get_LeftMargin() - pageSetup->get_RightMargin(), 0.01);
    
    pageSetup->set_Gutter(100.0);
    
    // Set the "RtlGutter" property to "true" to place the gutter in a more suitable position for right-to-left text.
    pageSetup->set_RtlGutter(true);
    
    // Set the "MultiplePages" property to "MultiplePagesType.MirrorMargins" to alternate
    // the left/right page side position of margins every page.
    pageSetup->set_MultiplePages(Aspose::Words::Settings::MultiplePagesType::MirrorMargins);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.Gutter.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.Gutter.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASPOSE_ASSERT_EQ(100.0, pageSetup->get_Gutter());
    ASSERT_TRUE(pageSetup->get_RtlGutter());
    ASSERT_EQ(Aspose::Words::Settings::MultiplePagesType::MirrorMargins, pageSetup->get_MultiplePages());
}

namespace gtest_test
{

TEST_F(ExPageSetup, Gutter)
{
    s_instance->Gutter();
}

} // namespace gtest_test

void ExPageSetup::Booklet()
{
    //ExStart
    //ExFor:PageSetup.Gutter
    //ExFor:PageSetup.MultiplePages
    //ExFor:PageSetup.SheetsPerBooklet
    //ExFor:MultiplePagesType
    //ExSummary:Shows how to configure a document that can be printed as a book fold.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Insert text that spans 16 pages.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"My Booklet:");
    
    for (int32_t i = 0; i < 15; i++)
    {
        builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
        builder->Write(System::String::Format(u"Booklet face #{0}", i));
    }
    
    // Configure the first section's "PageSetup" property to print the document in the form of a book fold.
    // When we print this document on both sides, we can take the pages to stack them
    // and fold them all down the middle at once. The contents of the document will line up into a book fold.
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->set_MultiplePages(Aspose::Words::Settings::MultiplePagesType::BookFoldPrinting);
    
    // We can only specify the number of sheets in multiples of 4.
    pageSetup->set_SheetsPerBooklet(4);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.Booklet.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.Booklet.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_EQ(Aspose::Words::Settings::MultiplePagesType::BookFoldPrinting, pageSetup->get_MultiplePages());
    ASSERT_EQ(4, pageSetup->get_SheetsPerBooklet());
}

namespace gtest_test
{

TEST_F(ExPageSetup, Booklet)
{
    s_instance->Booklet();
}

} // namespace gtest_test

void ExPageSetup::SetTextOrientation()
{
    //ExStart
    //ExFor:PageSetup.TextOrientation
    //ExSummary:Shows how to set text orientation.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    // Set the "TextOrientation" property to "TextOrientation.Upward" to rotate all the text 90 degrees
    // to the right so that all left-to-right text now goes top-to-bottom.
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->set_TextOrientation(Aspose::Words::TextOrientation::Upward);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.SetTextOrientation.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.SetTextOrientation.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_EQ(Aspose::Words::TextOrientation::Upward, pageSetup->get_TextOrientation());
}

namespace gtest_test
{

TEST_F(ExPageSetup, SetTextOrientation)
{
    s_instance->SetTextOrientation();
}

} // namespace gtest_test

void ExPageSetup::SuppressEndnotes()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->RemoveAllChildren();
    
    // By default, a document compiles all endnotes at its end. 
    ASSERT_EQ(Aspose::Words::Notes::EndnotePosition::EndOfDocument, doc->get_EndnoteOptions()->get_Position());
    
    // We use the "Position" property of the document's "EndnoteOptions" object
    // to collect endnotes at the end of each section instead. 
    doc->get_EndnoteOptions()->set_Position(Aspose::Words::Notes::EndnotePosition::EndOfSection);
    
    InsertSectionWithEndnote(doc, u"Section 1", u"Endnote 1, will stay in section 1");
    InsertSectionWithEndnote(doc, u"Section 2", u"Endnote 2, will be pushed down to section 3");
    InsertSectionWithEndnote(doc, u"Section 3", u"Endnote 3, will stay in section 3");
    
    // While getting sections to display their respective endnotes, we can set the "SuppressEndnotes" flag
    // of a section's "PageSetup" object to "true" to revert to the default behavior and pass its endnotes
    // onto the next section.
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_Sections()->idx_get(1)->get_PageSetup();
    pageSetup->set_SuppressEndnotes(true);
    
    doc->Save(get_ArtifactsDir() + u"PageSetup.SuppressEndnotes.docx");
    TestSuppressEndnotes(System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"PageSetup.SuppressEndnotes.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExPageSetup, SuppressEndnotes)
{
    s_instance->SuppressEndnotes();
}

} // namespace gtest_test

void ExPageSetup::ChapterPageSeparator()
{
    //ExStart
    //ExFor:PageSetup.HeadingLevelForChapter
    //ExFor:ChapterPageSeparator
    //ExFor:PageSetup.ChapterPageSeparator
    //ExSummary:Shows how to work with page chapters.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    pageSetup->set_PageNumberStyle(Aspose::Words::NumberStyle::UppercaseRoman);
    pageSetup->set_ChapterPageSeparator(Aspose::Words::ChapterPageSeparator::Colon);
    pageSetup->set_HeadingLevelForChapter(1);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPageSetup, ChapterPageSeparator)
{
    s_instance->ChapterPageSeparator();
}

} // namespace gtest_test

void ExPageSetup::JisbPaperSize()
{
    //ExStart:JisbPaperSize
    //GistId:12a3a3cfe30f3145220db88428a9f814
    //ExFor:PageSetup.PaperSize
    //ExSummary:Shows how to set the paper size of JisB4 or JisB5.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_FirstSection()->get_PageSetup();
    // Set the paper size to JisB4 (257x364mm).
    pageSetup->set_PaperSize(Aspose::Words::PaperSize::JisB4);
    // Alternatively, set the paper size to JisB5. (182x257mm).
    pageSetup->set_PaperSize(Aspose::Words::PaperSize::JisB5);
    //ExEnd:JisbPaperSize
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_EQ(Aspose::Words::PaperSize::JisB5, pageSetup->get_PaperSize());
}

namespace gtest_test
{

TEST_F(ExPageSetup, JisbPaperSize)
{
    s_instance->JisbPaperSize();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
