// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExPageSetup.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/enumerator_adapter.h>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/TextOrientation.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
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
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderType.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Settings;
namespace ApiExamples {

void ExPageSetup::InsertSection(SharedPtr<Aspose::Words::Document> doc, String sectionBodyText, String endnoteText)
{
    auto section = MakeObject<Section>(doc);

    doc->AppendChild(section);

    auto body = MakeObject<Body>(doc);
    section->AppendChild(body);

    ASPOSE_ASSERT_EQ(section, body->get_ParentNode());

    auto para = MakeObject<Paragraph>(doc);
    body->AppendChild(para);

    ASPOSE_ASSERT_EQ(body, para->get_ParentNode());

    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->MoveTo(para);
    builder->Write(sectionBodyText);
    builder->InsertFootnote(Aspose::Words::FootnoteType::Endnote, endnoteText);
}

void ExPageSetup::TestSuppressEndnotes(SharedPtr<Aspose::Words::Document> doc)
{
    SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(1)->get_PageSetup();

    ASSERT_TRUE(pageSetup->get_SuppressEndnotes());
}

namespace gtest_test
{

class ExPageSetup : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExPageSetup> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExPageSetup>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExPageSetup> ExPageSetup::s_instance;

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
    //ExSummary:Shows how to insert sections using DocumentBuilder, specify page setup for a section and reset page setup to defaults.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Modify the first section in the document
    builder->get_PageSetup()->set_Orientation(Aspose::Words::Orientation::Landscape);
    builder->get_PageSetup()->set_VerticalAlignment(Aspose::Words::PageVerticalAlignment::Center);
    builder->Writeln(u"Section 1, landscape oriented and text vertically centered.");

    // Start a new section and reset its formatting to defaults
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->get_PageSetup()->ClearFormatting();
    builder->Writeln(u"Section 2, back to default Letter paper size, portrait orientation and top alignment.");

    doc->Save(ArtifactsDir + u"PageSetup.ClearFormatting.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.ClearFormatting.docx");

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

void ExPageSetup::DifferentHeaders()
{
    //ExStart
    //ExFor:PageSetup.DifferentFirstPageHeaderFooter
    //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
    //ExFor:PageSetup.LayoutMode
    //ExFor:PageSetup.CharactersPerLine
    //ExFor:PageSetup.LinesPerPage
    //ExFor:SectionLayoutMode
    //ExSummary:Shows how to create headers and footers different for first, even and odd pages using DocumentBuilder.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<PageSetup> pageSetup = builder->get_PageSetup();
    pageSetup->set_DifferentFirstPageHeaderFooter(true);
    pageSetup->set_OddAndEvenPagesHeaderFooter(true);
    pageSetup->set_LayoutMode(Aspose::Words::SectionLayoutMode::LineGrid);
    pageSetup->set_CharactersPerLine(1);
    pageSetup->set_LinesPerPage(1);

    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderFirst);
    builder->Writeln(u"First page header.");

    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderEven);
    builder->Writeln(u"Even pages header.");

    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Writeln(u"Odd pages header.");

    // Move back to the main story of the first section
    builder->MoveToSection(0);
    builder->Writeln(u"Text page 1.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Text page 2.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Text page 3.");

    doc->Save(ArtifactsDir + u"PageSetup.DifferentHeaders.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.DifferentHeaders.docx");

    ASSERT_TRUE(pageSetup->get_DifferentFirstPageHeaderFooter());
    ASSERT_TRUE(pageSetup->get_OddAndEvenPagesHeaderFooter());
    ASSERT_EQ(Aspose::Words::SectionLayoutMode::LineGrid, doc->get_FirstSection()->get_PageSetup()->get_LayoutMode());
    ASSERT_EQ(1, doc->get_FirstSection()->get_PageSetup()->get_CharactersPerLine());
    ASSERT_EQ(1, doc->get_FirstSection()->get_PageSetup()->get_LinesPerPage());
}

namespace gtest_test
{

TEST_F(ExPageSetup, DifferentHeaders)
{
    s_instance->DifferentHeaders();
}

} // namespace gtest_test

void ExPageSetup::SetSectionStart()
{
    //ExStart
    //ExFor:SectionStart
    //ExFor:PageSetup.SectionStart
    //ExFor:Document.Sections
    //ExSummary:Shows how to specify how the section starts, from a new page, on the same page or other.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Add text to the first section and that comes with a blank document,
    // then add a new section that starts a new page and give it text as well
    builder->Writeln(u"This text is in section 1.");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Writeln(u"This text is in section 2.");

    // Section break types determine how a new section gets split from the previous section
    // By inserting a "SectionBreakNewPage" type section break, we've set this section's SectionStart value to "NewPage"
    ASSERT_EQ(Aspose::Words::SectionStart::NewPage, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_SectionStart());

    // Insert a new column section the same way
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewColumn);
    builder->Writeln(u"This text is in section 3.");

    ASSERT_EQ(Aspose::Words::SectionStart::NewColumn, doc->get_Sections()->idx_get(2)->get_PageSetup()->get_SectionStart());

    // We can change the types of section breaks by assigning different values to each section's SectionStart
    // Setting their values to "Continuous" will put no visible breaks between sections
    // and will leave all the content of this document on one page
    doc->get_Sections()->idx_get(1)->get_PageSetup()->set_SectionStart(Aspose::Words::SectionStart::Continuous);
    doc->get_Sections()->idx_get(2)->get_PageSetup()->set_SectionStart(Aspose::Words::SectionStart::Continuous);

    doc->Save(ArtifactsDir + u"PageSetup.SetSectionStart.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.SetSectionStart.docx");

    ASSERT_EQ(Aspose::Words::SectionStart::NewPage, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_SectionStart());
    ASSERT_EQ(Aspose::Words::SectionStart::Continuous, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_SectionStart());
    ASSERT_EQ(Aspose::Words::SectionStart::Continuous, doc->get_Sections()->idx_get(2)->get_PageSetup()->get_SectionStart());
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
    //ExSummary:Shows how to adjust paper size, orientation, margins and other settings for a section.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->get_PageSetup()->set_PaperSize(Aspose::Words::PaperSize::Legal);
    builder->get_PageSetup()->set_Orientation(Aspose::Words::Orientation::Landscape);
    builder->get_PageSetup()->set_TopMargin(ConvertUtil::InchToPoint(1.0));
    builder->get_PageSetup()->set_BottomMargin(ConvertUtil::InchToPoint(1.0));
    builder->get_PageSetup()->set_LeftMargin(ConvertUtil::InchToPoint(1.5));
    builder->get_PageSetup()->set_RightMargin(ConvertUtil::InchToPoint(1.5));
    builder->get_PageSetup()->set_HeaderDistance(ConvertUtil::InchToPoint(0.2));
    builder->get_PageSetup()->set_FooterDistance(ConvertUtil::InchToPoint(0.2));

    builder->Writeln(u"Hello world.");

    doc->Save(ArtifactsDir + u"PageSetup.PageMargins.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.PageMargins.docx");

    ASSERT_EQ(Aspose::Words::PaperSize::Legal, doc->get_FirstSection()->get_PageSetup()->get_PaperSize());
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

void ExPageSetup::ColumnsSameWidth()
{
    //ExStart
    //ExFor:PageSetup.TextColumns
    //ExFor:TextColumnCollection
    //ExFor:TextColumnCollection.Spacing
    //ExFor:TextColumnCollection.SetCount
    //ExSummary:Shows how to create multiple evenly spaced columns in a section using DocumentBuilder.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<TextColumnCollection> columns = builder->get_PageSetup()->get_TextColumns();
    // Make spacing between columns wider
    columns->set_Spacing(100);
    // This creates two columns of equal width
    columns->SetCount(2);

    builder->Writeln(u"Text in column 1.");
    builder->InsertBreak(Aspose::Words::BreakType::ColumnBreak);
    builder->Writeln(u"Text in column 2.");

    doc->Save(ArtifactsDir + u"PageSetup.ColumnsSameWidth.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.ColumnsSameWidth.docx");

    ASPOSE_ASSERT_EQ(100.0, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_Spacing());
    ASSERT_EQ(2, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_Count());
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
    //ExFor:TextColumnCollection.LineBetween
    //ExFor:TextColumnCollection.EvenlySpaced
    //ExFor:TextColumnCollection.Item
    //ExFor:TextColumn
    //ExFor:TextColumn.Width
    //ExFor:TextColumn.SpaceAfter
    //ExSummary:Shows how to set widths of columns.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<TextColumnCollection> columns = builder->get_PageSetup()->get_TextColumns();
    // Show vertical line between columns
    columns->set_LineBetween(true);
    // Indicate we want to create column with different widths
    columns->set_EvenlySpaced(false);
    // Create two columns, note they will be created with zero widths, need to set them
    columns->SetCount(2);

    // Set the first column to be narrow
    SharedPtr<TextColumn> column = columns->idx_get(0);
    column->set_Width(100);
    column->set_SpaceAfter(20);

    // Set the second column to take the rest of the space available on the page
    column = columns->idx_get(1);
    SharedPtr<PageSetup> pageSetup = builder->get_PageSetup();
    double contentWidth = pageSetup->get_PageWidth() - pageSetup->get_LeftMargin() - pageSetup->get_RightMargin();
    column->set_Width(contentWidth - column->get_Width() - column->get_SpaceAfter());

    builder->Writeln(u"Narrow column 1.");
    builder->InsertBreak(Aspose::Words::BreakType::ColumnBreak);
    builder->Writeln(u"Wide column 2.");

    doc->Save(ArtifactsDir + u"PageSetup.CustomColumnWidth.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.CustomColumnWidth.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();

    ASSERT_TRUE(pageSetup->get_TextColumns()->get_LineBetween());
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

void ExPageSetup::LineNumbers()
{
    //ExStart
    //ExFor:PageSetup.LineStartingNumber
    //ExFor:PageSetup.LineNumberDistanceFromText
    //ExFor:PageSetup.LineNumberCountBy
    //ExFor:PageSetup.LineNumberRestartMode
    //ExFor:ParagraphFormat.SuppressLineNumbers
    //ExFor:LineNumberRestartMode
    //ExSummary:Shows how to enable Microsoft Word line numbering for a section.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Line numbering for each section can be configured via PageSetup
    SharedPtr<PageSetup> pageSetup = builder->get_PageSetup();
    pageSetup->set_LineStartingNumber(1);
    pageSetup->set_LineNumberCountBy(3);
    pageSetup->set_LineNumberRestartMode(Aspose::Words::LineNumberRestartMode::RestartPage);
    pageSetup->set_LineNumberDistanceFromText(50.0);

    // LineNumberCountBy is set to 3, so every line that's a multiple of 3
    // will display that line number to the left of the text
    for (int i = 1; i <= 25; i++)
    {
        builder->Writeln(String::Format(u"Line {0}.",i));
    }

    // The line counter will skip any paragraph with this flag set to true
    // Normally, the number "15" would normally appear next to this paragraph, which says "Line 15"
    // Since we set this flag to true and this paragraph is not counted by numbering,
    // number 15 will appear next to the next paragraph, "Line 16", and from then on counting will carry on as normal
    // until it will restart according to LineNumberRestartMode
    doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(14)->get_ParagraphFormat()->set_SuppressLineNumbers(true);

    doc->Save(ArtifactsDir + u"PageSetup.LineNumbers.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.LineNumbers.docx");
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
    //ExSummary:Shows how to create a page border that looks like a wide blue band at the top of the first page only.
    auto doc = MakeObject<Document>();

    SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->set_BorderAlwaysInFront(false);
    pageSetup->set_BorderDistanceFrom(Aspose::Words::PageBorderDistanceFrom::PageEdge);
    pageSetup->set_BorderAppliesTo(Aspose::Words::PageBorderAppliesTo::FirstPage);

    SharedPtr<Border> border = pageSetup->get_Borders()->idx_get(Aspose::Words::BorderType::Top);
    border->set_LineStyle(Aspose::Words::LineStyle::Single);
    border->set_LineWidth(30);
    border->set_Color(System::Drawing::Color::get_Blue());
    border->set_DistanceFromText(0);

    doc->Save(ArtifactsDir + u"PageSetup.PageBorderProperties.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.PageBorderProperties.docx");
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
    auto doc = MakeObject<Document>();
    SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();

    pageSetup->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::DoubleWave);
    pageSetup->get_Borders()->set_LineWidth(2);
    pageSetup->get_Borders()->set_Color(System::Drawing::Color::get_Green());
    pageSetup->get_Borders()->set_DistanceFromText(24);
    pageSetup->get_Borders()->set_Shadow(true);

    doc->Save(ArtifactsDir + u"PageSetup.PageBorders.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.PageBorders.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();

    for (auto border : System::IterateOver(pageSetup->get_Borders()))
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
    //ExSummary:Shows how to control page numbering per section.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Writeln(u"Section 1");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Writeln(u"Section 2");

    // Use document builder to create a header with a page number field for the first section
    // The page number will look like "Page V"
    builder->MoveToSection(0);
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Write(u"Page ");
    builder->InsertField(u"PAGE", u"");

    // Set first section page numbering
    SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->set_RestartPageNumbering(true);
    pageSetup->set_PageStartingNumber(5);
    pageSetup->set_PageNumberStyle(Aspose::Words::NumberStyle::UppercaseRoman);

    // Create a header for the section
    // The page number will look like " - 10 - ".
    builder->MoveToSection(1);
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    builder->Write(u" - ");
    builder->InsertField(u"PAGE", u"");
    builder->Write(u" - ");

    // Set second section page numbering
    pageSetup = doc->get_Sections()->idx_get(1)->get_PageSetup();
    pageSetup->set_PageStartingNumber(10);
    pageSetup->set_RestartPageNumbering(true);
    pageSetup->set_PageNumberStyle(Aspose::Words::NumberStyle::Arabic);

    doc->Save(ArtifactsDir + u"PageSetup.PageNumbering.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.PageNumbering.docx");
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
    //ExSummary:Shows how to set options for footnotes and endnotes in current section.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert text and a reference for it in the form of a footnote
    builder->Write(u"Hello world!.");
    builder->InsertFootnote(Aspose::Words::FootnoteType::Footnote, u"Footnote reference text.");

    // Set options for footnote position and numbering
    SharedPtr<Aspose::Words::FootnoteOptions> footnoteOptions = doc->get_Sections()->idx_get(0)->get_PageSetup()->get_FootnoteOptions();
    footnoteOptions->set_Position(Aspose::Words::FootnotePosition::BeneathText);
    footnoteOptions->set_RestartRule(Aspose::Words::FootnoteNumberingRule::RestartPage);
    footnoteOptions->set_StartNumber(1);

    // Endnotes also have a similar options object
    builder->Write(u" Hello again.");
    builder->InsertFootnote(Aspose::Words::FootnoteType::Footnote, u"Endnote reference text.");

    SharedPtr<EndnoteOptions> endnoteOptions = doc->get_Sections()->idx_get(0)->get_PageSetup()->get_EndnoteOptions();
    endnoteOptions->set_Position(Aspose::Words::EndnotePosition::EndOfDocument);
    endnoteOptions->set_RestartRule(Aspose::Words::FootnoteNumberingRule::Continuous);
    endnoteOptions->set_StartNumber(1);

    doc->Save(ArtifactsDir + u"PageSetup.FootnoteOptions.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.FootnoteOptions.docx");
    footnoteOptions = doc->get_FirstSection()->get_PageSetup()->get_FootnoteOptions();

    ASSERT_EQ(Aspose::Words::FootnotePosition::BeneathText, footnoteOptions->get_Position());
    ASSERT_EQ(Aspose::Words::FootnoteNumberingRule::RestartPage, footnoteOptions->get_RestartRule());
    ASSERT_EQ(1, footnoteOptions->get_StartNumber());

    endnoteOptions = doc->get_FirstSection()->get_PageSetup()->get_EndnoteOptions();

    ASSERT_EQ(Aspose::Words::EndnotePosition::EndOfDocument, endnoteOptions->get_Position());
    ASSERT_EQ(Aspose::Words::FootnoteNumberingRule::Continuous, endnoteOptions->get_RestartRule());
    ASSERT_EQ(1, endnoteOptions->get_StartNumber());
}

namespace gtest_test
{

TEST_F(ExPageSetup, FootnoteOptions)
{
    s_instance->FootnoteOptions();
}

} // namespace gtest_test

void ExPageSetup::Bidi()
{
    //ExStart
    //ExFor:PageSetup.Bidi
    //ExSummary:Shows how to change the order of columns.
    auto doc = MakeObject<Document>();

    SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->get_TextColumns()->SetCount(3);

    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Write(u"Column 1.");
    builder->InsertBreak(Aspose::Words::BreakType::ColumnBreak);
    builder->Write(u"Column 2.");
    builder->InsertBreak(Aspose::Words::BreakType::ColumnBreak);
    builder->Write(u"Column 3.");

    // Reverse the order of the columns
    pageSetup->set_Bidi(true);

    doc->Save(ArtifactsDir + u"PageSetup.Bidi.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.Bidi.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();

    ASSERT_EQ(3, pageSetup->get_TextColumns()->get_Count());
    ASSERT_TRUE(pageSetup->get_Bidi());
}

namespace gtest_test
{

TEST_F(ExPageSetup, Bidi)
{
    s_instance->Bidi();
}

} // namespace gtest_test

void ExPageSetup::PageBorder()
{
    //ExStart
    //ExFor:PageSetup.BorderSurroundsFooter
    //ExFor:PageSetup.BorderSurroundsHeader
    //ExSummary:Shows how to apply a border to the page and header/footer.
    auto doc = MakeObject<Document>();

    // Insert header and footer text
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Write(u"Header");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterPrimary);
    builder->Write(u"Footer");
    builder->MoveToDocumentEnd();

    // Insert a page border and set the color and line style
    SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::Double);
    pageSetup->get_Borders()->set_Color(System::Drawing::Color::get_Blue());

    // By default, page borders don't surround headers and footers
    // We can change that by setting these flags
    pageSetup->set_BorderSurroundsFooter(true);
    pageSetup->set_BorderSurroundsHeader(true);

    doc->Save(ArtifactsDir + u"PageSetup.PageBorder.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.PageBorder.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();

    ASSERT_TRUE(pageSetup->get_BorderSurroundsFooter());
    ASSERT_TRUE(pageSetup->get_BorderSurroundsHeader());
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
    auto doc = MakeObject<Document>();

    // Insert text spanning several pages
    auto builder = MakeObject<DocumentBuilder>(doc);
    for (int i = 0; i < 6; i++)
    {
        builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    }

    // We can access the gutter margin in the section's page options,
    // which is a margin which is added to the page margin at one side of the page
    SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->set_Gutter(100.0);

    // If our text is LTR, the gutter will appear on the left side of the page
    // Setting this flag will move it to the right side
    pageSetup->set_RtlGutter(true);

    // Mirroring the margins will make the gutter alternate in position from page to page
    pageSetup->set_MultiplePages(Aspose::Words::Settings::MultiplePagesType::MirrorMargins);

    doc->Save(ArtifactsDir + u"PageSetup.Gutter.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.Gutter.docx");
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
    //ExFor:PageSetup.SheetsPerBooklet
    //ExSummary:Shows how to create a booklet.
    auto doc = MakeObject<Document>();

    // Use a document builder to create 16 pages of content that will be compiled in a booklet
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"My Booklet:");

    for (int i = 0; i < 15; i++)
    {
        builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
        builder->Write(String::Format(u"Booklet face #{0}",i));
    }

    // Set the number of sheets that will be used by the printer to create the booklet
    // After being printed on both sides, the sheets can be stacked and folded down the centre
    // The contents that we placed in such a way that they will be in order once the booklet is folded
    // We can only specify the number of sheets in multiples of 4
    SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->set_MultiplePages(Aspose::Words::Settings::MultiplePagesType::BookFoldPrinting);
    pageSetup->set_SheetsPerBooklet(4);

    doc->Save(ArtifactsDir + u"PageSetup.Booklet.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.Booklet.docx");
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

void ExPageSetup::SectionTextOrientation()
{
    //ExStart
    //ExFor:PageSetup.TextOrientation
    //ExSummary:Shows how to set text orientation.
    auto doc = MakeObject<Document>();

    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");

    // Setting this value will rotate the section's text 90 degrees to the right
    SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->set_TextOrientation(Aspose::Words::TextOrientation::Upward);

    doc->Save(ArtifactsDir + u"PageSetup.SectionTextOrientation.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.SectionTextOrientation.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();

    ASSERT_EQ(Aspose::Words::TextOrientation::Upward, pageSetup->get_TextOrientation());
}

namespace gtest_test
{

TEST_F(ExPageSetup, SectionTextOrientation)
{
    s_instance->SectionTextOrientation();
}

} // namespace gtest_test

void ExPageSetup::SuppressEndnotes()
{
    // Create a new document and make it empty
    auto doc = MakeObject<Document>();
    doc->RemoveAllChildren();

    // Normally endnotes are all stored at the end of a document, but this option lets us store them at the end of each section
    doc->get_EndnoteOptions()->set_Position(Aspose::Words::EndnotePosition::EndOfSection);

    // Create 3 new sections, each having a paragraph and an endnote at the end
    InsertSection(doc, u"Section 1", u"Endnote 1, will stay in section 1");
    InsertSection(doc, u"Section 2", u"Endnote 2, will be pushed down to section 3");
    InsertSection(doc, u"Section 3", u"Endnote 3, will stay in section 3");

    // Each section contains its own page setup object
    // Setting this value will push this section's endnotes down to the next section
    SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(1)->get_PageSetup();
    pageSetup->set_SuppressEndnotes(true);

    doc->Save(ArtifactsDir + u"PageSetup.SuppressEndnotes.docx");
    TestSuppressEndnotes(MakeObject<Document>(ArtifactsDir + u"PageSetup.SuppressEndnotes.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExPageSetup, SuppressEndnotes)
{
    s_instance->SuppressEndnotes();
}

} // namespace gtest_test

} // namespace ApiExamples
