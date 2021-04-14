#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderType.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Footnotes/EndnoteOptions.h>
#include <Aspose.Words.Cpp/Model/Footnotes/EndnotePosition.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteNumberingRule.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteOptions.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnotePosition.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Numbering/NumberStyle.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/LineNumberRestartMode.h>
#include <Aspose.Words.Cpp/Model/Sections/Orientation.h>
#include <Aspose.Words.Cpp/Model/Sections/PageBorderAppliesTo.h>
#include <Aspose.Words.Cpp/Model/Sections/PageBorderDistanceFrom.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/PageVerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Sections/PaperSize.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionLayoutMode.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionStart.h>
#include <Aspose.Words.Cpp/Model/Sections/TextColumn.h>
#include <Aspose.Words.Cpp/Model/Sections/TextColumnCollection.h>
#include <Aspose.Words.Cpp/Model/Settings/MultiplePagesType.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/TextOrientation.h>
#include <drawing/color.h>
#include <system/enumerator_adapter.h>
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
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Settings;

namespace ApiExamples {

class ExPageSetup : public ApiExampleBase
{
public:
    void ClearFormatting()
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
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Modify the page setup properties for the builder's current section and add text.
        builder->get_PageSetup()->set_Orientation(Orientation::Landscape);
        builder->get_PageSetup()->set_VerticalAlignment(PageVerticalAlignment::Center);
        builder->Writeln(u"This is the first section, which landscape oriented with vertically centered text.");

        // If we start a new section using a document builder,
        // it will inherit the builder's current page setup properties.
        builder->InsertBreak(BreakType::SectionBreakNewPage);

        ASSERT_EQ(Orientation::Landscape, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_Orientation());
        ASSERT_EQ(PageVerticalAlignment::Center, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_VerticalAlignment());

        // We can revert its page setup properties to their default values using the "ClearFormatting" method.
        builder->get_PageSetup()->ClearFormatting();

        ASSERT_EQ(Orientation::Portrait, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_Orientation());
        ASSERT_EQ(PageVerticalAlignment::Top, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_VerticalAlignment());

        builder->Writeln(u"This is the second section, which is in default Letter paper size, portrait orientation and top alignment.");

        doc->Save(ArtifactsDir + u"PageSetup.ClearFormatting.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.ClearFormatting.docx");

        ASSERT_EQ(Orientation::Landscape, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_Orientation());
        ASSERT_EQ(PageVerticalAlignment::Center, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_VerticalAlignment());

        ASSERT_EQ(Orientation::Portrait, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_Orientation());
        ASSERT_EQ(PageVerticalAlignment::Top, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_VerticalAlignment());
    }

    void DifferentFirstPageHeaderFooter(bool differentFirstPageHeaderFooter)
    {
        //ExStart
        //ExFor:PageSetup.DifferentFirstPageHeaderFooter
        //ExSummary:Shows how to enable or disable primary headers/footers.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are two types of header/footers.
        // 1 -  The "First" header/footer, which appears on the first page of the section.
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderFirst);
        builder->Writeln(u"First page header.");

        builder->MoveToHeaderFooter(HeaderFooterType::FooterFirst);
        builder->Writeln(u"First page footer.");

        // 2 -  The "Primary" header/footer, which appears on every page in the section.
        // We can override the primary header/footer by a first and an even page header/footer.
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Writeln(u"Primary header.");

        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->Writeln(u"Primary footer.");

        builder->MoveToSection(0);
        builder->Writeln(u"Page 1.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 2.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 3.");

        // Each section has a "PageSetup" object that specifies page appearance-related properties
        // such as orientation, size, and borders.
        // Set the "DifferentFirstPageHeaderFooter" property to "true" to apply the first header/footer to the first page.
        // Set the "DifferentFirstPageHeaderFooter" property to "false"
        // to make the first page display the primary header/footer.
        builder->get_PageSetup()->set_DifferentFirstPageHeaderFooter(differentFirstPageHeaderFooter);

        doc->Save(ArtifactsDir + u"PageSetup.DifferentFirstPageHeaderFooter.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.DifferentFirstPageHeaderFooter.docx");

        ASPOSE_ASSERT_EQ(differentFirstPageHeaderFooter, doc->get_FirstSection()->get_PageSetup()->get_DifferentFirstPageHeaderFooter());
    }

    void OddAndEvenPagesHeaderFooter(bool oddAndEvenPagesHeaderFooter)
    {
        //ExStart
        //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
        //ExSummary:Shows how to enable or disable even page headers/footers.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are two types of header/footers.
        // 1 -  The "Primary" header/footer, which appears on every page in the section.
        // We can override the primary header/footer by a first and an even page header/footer.
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Writeln(u"Primary header.");

        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->Writeln(u"Primary footer.");

        // 2 -  The "Even" header/footer, which appears on every even page of this section.
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderEven);
        builder->Writeln(u"Even page header.");

        builder->MoveToHeaderFooter(HeaderFooterType::FooterEven);
        builder->Writeln(u"Even page footer.");

        builder->MoveToSection(0);
        builder->Writeln(u"Page 1.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 2.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 3.");

        // Each section has a "PageSetup" object that specifies page appearance-related properties
        // such as orientation, size, and borders.
        // Set the "OddAndEvenPagesHeaderFooter" property to "true"
        // to display the even page header/footer on even pages.
        // Set the "OddAndEvenPagesHeaderFooter" property to "false"
        // to display the primary header/footer on even pages.
        builder->get_PageSetup()->set_OddAndEvenPagesHeaderFooter(oddAndEvenPagesHeaderFooter);

        doc->Save(ArtifactsDir + u"PageSetup.OddAndEvenPagesHeaderFooter.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.OddAndEvenPagesHeaderFooter.docx");

        ASPOSE_ASSERT_EQ(oddAndEvenPagesHeaderFooter, doc->get_FirstSection()->get_PageSetup()->get_OddAndEvenPagesHeaderFooter());
    }

    void CharactersPerLine()
    {
        //ExStart
        //ExFor:PageSetup.CharactersPerLine
        //ExFor:PageSetup.LayoutMode
        //ExFor:SectionLayoutMode
        //ExSummary:Shows how to specify a for the number of characters that each line may have.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Enable pitching, and then use it to set the number of characters per line in this section.
        builder->get_PageSetup()->set_LayoutMode(SectionLayoutMode::Grid);
        builder->get_PageSetup()->set_CharactersPerLine(10);

        // The number of characters also depends on the size of the font.
        doc->get_Styles()->idx_get(u"Normal")->get_Font()->set_Size(20);

        ASSERT_EQ(8, doc->get_FirstSection()->get_PageSetup()->get_CharactersPerLine());

        builder->Writeln(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        doc->Save(ArtifactsDir + u"PageSetup.CharactersPerLine.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.CharactersPerLine.docx");

        ASSERT_EQ(SectionLayoutMode::Grid, doc->get_FirstSection()->get_PageSetup()->get_LayoutMode());
        ASSERT_EQ(8, doc->get_FirstSection()->get_PageSetup()->get_CharactersPerLine());
    }

    void LinesPerPage()
    {
        //ExStart
        //ExFor:PageSetup.LinesPerPage
        //ExFor:PageSetup.LayoutMode
        //ExFor:ParagraphFormat.SnapToGrid
        //ExFor:SectionLayoutMode
        //ExSummary:Shows how to specify a limit for the number of lines that each page may have.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Enable pitching, and then use it to set the number of lines per page in this section.
        // A large enough font size will push some lines down onto the next page to avoid overlapping characters.
        builder->get_PageSetup()->set_LayoutMode(SectionLayoutMode::LineGrid);
        builder->get_PageSetup()->set_LinesPerPage(15);

        builder->get_ParagraphFormat()->set_SnapToGrid(true);

        for (int i = 0; i < 30; i++)
        {
            builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");
        }

        doc->Save(ArtifactsDir + u"PageSetup.LinesPerPage.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.LinesPerPage.docx");

        ASSERT_EQ(SectionLayoutMode::LineGrid, doc->get_FirstSection()->get_PageSetup()->get_LayoutMode());
        ASSERT_EQ(15, doc->get_FirstSection()->get_PageSetup()->get_LinesPerPage());

        for (auto paragraph : System::IterateOver<Paragraph>(doc->get_FirstSection()->get_Body()->get_Paragraphs()))
        {
            ASSERT_TRUE(paragraph->get_ParagraphFormat()->get_SnapToGrid());
        }
    }

    void SetSectionStart()
    {
        //ExStart
        //ExFor:SectionStart
        //ExFor:PageSetup.SectionStart
        //ExFor:Document.Sections
        //ExSummary:Shows how to specify how a new section separates itself from the previous.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"This text is in section 1.");

        // Section break types determine how a new section separates itself from the previous section.
        // Below are five types of section breaks.
        // 1 -  Starts the next section on a new page:
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Writeln(u"This text is in section 2.");

        ASSERT_EQ(SectionStart::NewPage, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_SectionStart());

        // 2 -  Starts the next section on the current page:
        builder->InsertBreak(BreakType::SectionBreakContinuous);
        builder->Writeln(u"This text is in section 3.");

        ASSERT_EQ(SectionStart::Continuous, doc->get_Sections()->idx_get(2)->get_PageSetup()->get_SectionStart());

        // 3 -  Starts the next section on a new even page:
        builder->InsertBreak(BreakType::SectionBreakEvenPage);
        builder->Writeln(u"This text is in section 4.");

        ASSERT_EQ(SectionStart::EvenPage, doc->get_Sections()->idx_get(3)->get_PageSetup()->get_SectionStart());

        // 4 -  Starts the next section on a new odd page:
        builder->InsertBreak(BreakType::SectionBreakOddPage);
        builder->Writeln(u"This text is in section 5.");

        ASSERT_EQ(SectionStart::OddPage, doc->get_Sections()->idx_get(4)->get_PageSetup()->get_SectionStart());

        // 5 -  Starts the next section on a new column:
        SharedPtr<TextColumnCollection> columns = builder->get_PageSetup()->get_TextColumns();
        columns->SetCount(2);

        builder->InsertBreak(BreakType::SectionBreakNewColumn);
        builder->Writeln(u"This text is in section 6.");

        ASSERT_EQ(SectionStart::NewColumn, doc->get_Sections()->idx_get(5)->get_PageSetup()->get_SectionStart());

        doc->Save(ArtifactsDir + u"PageSetup.SetSectionStart.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.SetSectionStart.docx");

        ASSERT_EQ(SectionStart::NewPage, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_SectionStart());
        ASSERT_EQ(SectionStart::NewPage, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_SectionStart());
        ASSERT_EQ(SectionStart::Continuous, doc->get_Sections()->idx_get(2)->get_PageSetup()->get_SectionStart());
        ASSERT_EQ(SectionStart::EvenPage, doc->get_Sections()->idx_get(3)->get_PageSetup()->get_SectionStart());
        ASSERT_EQ(SectionStart::OddPage, doc->get_Sections()->idx_get(4)->get_PageSetup()->get_SectionStart());
        ASSERT_EQ(SectionStart::NewColumn, doc->get_Sections()->idx_get(5)->get_PageSetup()->get_SectionStart());
    }

    void PageMargins()
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
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_PageSetup()->set_PaperSize(PaperSize::Legal);
        builder->get_PageSetup()->set_Orientation(Orientation::Landscape);
        builder->get_PageSetup()->set_TopMargin(ConvertUtil::InchToPoint(1.0));
        builder->get_PageSetup()->set_BottomMargin(ConvertUtil::InchToPoint(1.0));
        builder->get_PageSetup()->set_LeftMargin(ConvertUtil::InchToPoint(1.5));
        builder->get_PageSetup()->set_RightMargin(ConvertUtil::InchToPoint(1.5));
        builder->get_PageSetup()->set_HeaderDistance(ConvertUtil::InchToPoint(0.2));
        builder->get_PageSetup()->set_FooterDistance(ConvertUtil::InchToPoint(0.2));

        builder->Writeln(u"Hello world!");

        doc->Save(ArtifactsDir + u"PageSetup.PageMargins.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.PageMargins.docx");

        ASSERT_EQ(PaperSize::Legal, doc->get_FirstSection()->get_PageSetup()->get_PaperSize());
        ASPOSE_ASSERT_EQ(1008.0, doc->get_FirstSection()->get_PageSetup()->get_PageWidth());
        ASPOSE_ASSERT_EQ(612.0, doc->get_FirstSection()->get_PageSetup()->get_PageHeight());
        ASSERT_EQ(Orientation::Landscape, doc->get_FirstSection()->get_PageSetup()->get_Orientation());
        ASPOSE_ASSERT_EQ(72.0, doc->get_FirstSection()->get_PageSetup()->get_TopMargin());
        ASPOSE_ASSERT_EQ(72.0, doc->get_FirstSection()->get_PageSetup()->get_BottomMargin());
        ASPOSE_ASSERT_EQ(108.0, doc->get_FirstSection()->get_PageSetup()->get_LeftMargin());
        ASPOSE_ASSERT_EQ(108.0, doc->get_FirstSection()->get_PageSetup()->get_RightMargin());
        ASPOSE_ASSERT_EQ(14.4, doc->get_FirstSection()->get_PageSetup()->get_HeaderDistance());
        ASPOSE_ASSERT_EQ(14.4, doc->get_FirstSection()->get_PageSetup()->get_FooterDistance());
    }

    void PaperSizes()
    {
        //ExStart
        //ExFor:PaperSize
        //ExFor:PageSetup.PaperSize
        //ExSummary:Shows how to set page sizes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // We can change the current page's size to a pre-defined size
        // by using the "PaperSize" property of this section's PageSetup object.
        builder->get_PageSetup()->set_PaperSize(PaperSize::Tabloid);

        ASPOSE_ASSERT_EQ(792.0, builder->get_PageSetup()->get_PageWidth());
        ASPOSE_ASSERT_EQ(1224.0, builder->get_PageSetup()->get_PageHeight());

        builder->Writeln(String::Format(u"This page is {0}x{1}.", builder->get_PageSetup()->get_PageWidth(), builder->get_PageSetup()->get_PageHeight()));

        // Each section has its own PageSetup object. When we use a document builder to make a new section,
        // that section's PageSetup object inherits all the previous section's PageSetup object's values.
        builder->InsertBreak(BreakType::SectionBreakEvenPage);

        ASSERT_EQ(PaperSize::Tabloid, builder->get_PageSetup()->get_PaperSize());

        builder->get_PageSetup()->set_PaperSize(PaperSize::A5);
        builder->Writeln(String::Format(u"This page is {0}x{1}.", builder->get_PageSetup()->get_PageWidth(), builder->get_PageSetup()->get_PageHeight()));

        ASPOSE_ASSERT_EQ(419.55, builder->get_PageSetup()->get_PageWidth());
        ASPOSE_ASSERT_EQ(595.30, builder->get_PageSetup()->get_PageHeight());

        builder->InsertBreak(BreakType::SectionBreakEvenPage);

        // Set a custom size for this section's pages.
        builder->get_PageSetup()->set_PageWidth(620);
        builder->get_PageSetup()->set_PageHeight(480);

        ASSERT_EQ(PaperSize::Custom, builder->get_PageSetup()->get_PaperSize());

        builder->Writeln(String::Format(u"This page is {0}x{1}.", builder->get_PageSetup()->get_PageWidth(), builder->get_PageSetup()->get_PageHeight()));

        doc->Save(ArtifactsDir + u"PageSetup.PaperSizes.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.PaperSizes.docx");

        ASSERT_EQ(PaperSize::Tabloid, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_PaperSize());
        ASPOSE_ASSERT_EQ(792.0, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_PageWidth());
        ASPOSE_ASSERT_EQ(1224.0, doc->get_Sections()->idx_get(0)->get_PageSetup()->get_PageHeight());
        ASSERT_EQ(PaperSize::A5, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_PaperSize());
        ASPOSE_ASSERT_EQ(419.55, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_PageWidth());
        ASPOSE_ASSERT_EQ(595.30, doc->get_Sections()->idx_get(1)->get_PageSetup()->get_PageHeight());
        ASSERT_EQ(PaperSize::Custom, doc->get_Sections()->idx_get(2)->get_PageSetup()->get_PaperSize());
        ASPOSE_ASSERT_EQ(620.0, doc->get_Sections()->idx_get(2)->get_PageSetup()->get_PageWidth());
        ASPOSE_ASSERT_EQ(480.0, doc->get_Sections()->idx_get(2)->get_PageSetup()->get_PageHeight());
    }

    void ColumnsSameWidth()
    {
        //ExStart
        //ExFor:PageSetup.TextColumns
        //ExFor:TextColumnCollection
        //ExFor:TextColumnCollection.Spacing
        //ExFor:TextColumnCollection.SetCount
        //ExSummary:Shows how to create multiple evenly spaced columns in a section.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<TextColumnCollection> columns = builder->get_PageSetup()->get_TextColumns();
        columns->set_Spacing(100);
        columns->SetCount(2);

        builder->Writeln(u"Column 1.");
        builder->InsertBreak(BreakType::ColumnBreak);
        builder->Writeln(u"Column 2.");

        doc->Save(ArtifactsDir + u"PageSetup.ColumnsSameWidth.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.ColumnsSameWidth.docx");

        ASPOSE_ASSERT_EQ(100.0, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_Spacing());
        ASSERT_EQ(2, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_Count());
    }

    void CustomColumnWidth()
    {
        //ExStart
        //ExFor:TextColumnCollection.EvenlySpaced
        //ExFor:TextColumnCollection.Item
        //ExFor:TextColumn
        //ExFor:TextColumn.Width
        //ExFor:TextColumn.SpaceAfter
        //ExSummary:Shows how to create unevenly spaced columns.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        SharedPtr<PageSetup> pageSetup = builder->get_PageSetup();

        SharedPtr<TextColumnCollection> columns = pageSetup->get_TextColumns();
        columns->set_EvenlySpaced(false);
        columns->SetCount(2);

        // Determine the amount of room that we have available for arranging columns.
        double contentWidth = pageSetup->get_PageWidth() - pageSetup->get_LeftMargin() - pageSetup->get_RightMargin();

        ASSERT_NEAR(470.30, contentWidth, 0.01);

        // Set the first column to be narrow.
        SharedPtr<TextColumn> column = columns->idx_get(0);
        column->set_Width(100);
        column->set_SpaceAfter(20);

        // Set the second column to take the rest of the space available within the margins of the page.
        column = columns->idx_get(1);
        column->set_Width(contentWidth - column->get_Width() - column->get_SpaceAfter());

        builder->Writeln(u"Narrow column 1.");
        builder->InsertBreak(BreakType::ColumnBreak);
        builder->Writeln(u"Wide column 2.");

        doc->Save(ArtifactsDir + u"PageSetup.CustomColumnWidth.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.CustomColumnWidth.docx");
        pageSetup = doc->get_FirstSection()->get_PageSetup();

        ASSERT_FALSE(pageSetup->get_TextColumns()->get_EvenlySpaced());
        ASSERT_EQ(2, pageSetup->get_TextColumns()->get_Count());
        ASPOSE_ASSERT_EQ(100.0, pageSetup->get_TextColumns()->idx_get(0)->get_Width());
        ASPOSE_ASSERT_EQ(20.0, pageSetup->get_TextColumns()->idx_get(0)->get_SpaceAfter());
        ASPOSE_ASSERT_EQ(470.3, pageSetup->get_TextColumns()->idx_get(1)->get_Width());
        ASPOSE_ASSERT_EQ(0.0, pageSetup->get_TextColumns()->idx_get(1)->get_SpaceAfter());
    }

    void VerticalLineBetweenColumns(bool lineBetween)
    {
        //ExStart
        //ExFor:TextColumnCollection.LineBetween
        //ExSummary:Shows how to separate columns with a vertical line.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Configure the current section's PageSetup object to divide the text into several columns.
        // Set the "LineBetween" property to "true" to put a dividing line between columns.
        // Set the "LineBetween" property to "false" to leave the space between columns blank.
        SharedPtr<TextColumnCollection> columns = builder->get_PageSetup()->get_TextColumns();
        columns->set_LineBetween(lineBetween);
        columns->SetCount(3);

        builder->Writeln(u"Column 1.");
        builder->InsertBreak(BreakType::ColumnBreak);
        builder->Writeln(u"Column 2.");
        builder->InsertBreak(BreakType::ColumnBreak);
        builder->Writeln(u"Column 3.");

        doc->Save(ArtifactsDir + u"PageSetup.VerticalLineBetweenColumns.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.VerticalLineBetweenColumns.docx");

        ASPOSE_ASSERT_EQ(lineBetween, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_LineBetween());
    }

    void LineNumbers()
    {
        //ExStart
        //ExFor:PageSetup.LineStartingNumber
        //ExFor:PageSetup.LineNumberDistanceFromText
        //ExFor:PageSetup.LineNumberCountBy
        //ExFor:PageSetup.LineNumberRestartMode
        //ExFor:ParagraphFormat.SuppressLineNumbers
        //ExFor:LineNumberRestartMode
        //ExSummary:Shows how to enable line numbering for a section.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // We can use the section's PageSetup object to display numbers to the left of the section's text lines.
        // This is the same behavior as a List object,
        // but it covers the entire section and does not modify the text in any way.
        // Our section will restart the numbering on each new page from 1 and display the number,
        // if it is a multiple of 3, at 50pt to the left of the line.
        SharedPtr<PageSetup> pageSetup = builder->get_PageSetup();
        pageSetup->set_LineStartingNumber(1);
        pageSetup->set_LineNumberCountBy(3);
        pageSetup->set_LineNumberRestartMode(LineNumberRestartMode::RestartPage);
        pageSetup->set_LineNumberDistanceFromText(50.0);

        for (int i = 1; i <= 25; i++)
        {
            builder->Writeln(String::Format(u"Line {0}.", i));
        }

        // The line counter will skip any paragraph with the "SuppressLineNumbers" flag set to "true".
        // This paragraph is on the 15th line, which is a multiple of 3, and thus would normally display a line number.
        // The section's line counter will also ignore this line, treat the next line as the 15th,
        // and continue the count from that point onward.
        doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(14)->get_ParagraphFormat()->set_SuppressLineNumbers(true);

        doc->Save(ArtifactsDir + u"PageSetup.LineNumbers.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.LineNumbers.docx");
        pageSetup = doc->get_FirstSection()->get_PageSetup();

        ASSERT_EQ(1, pageSetup->get_LineStartingNumber());
        ASSERT_EQ(3, pageSetup->get_LineNumberCountBy());
        ASSERT_EQ(LineNumberRestartMode::RestartPage, pageSetup->get_LineNumberRestartMode());
        ASPOSE_ASSERT_EQ(50.0, pageSetup->get_LineNumberDistanceFromText());
    }

    void PageBorderProperties()
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
        auto doc = MakeObject<Document>();

        SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
        pageSetup->set_BorderAlwaysInFront(false);
        pageSetup->set_BorderDistanceFrom(PageBorderDistanceFrom::PageEdge);
        pageSetup->set_BorderAppliesTo(PageBorderAppliesTo::FirstPage);

        SharedPtr<Border> border = pageSetup->get_Borders()->idx_get(BorderType::Top);
        border->set_LineStyle(LineStyle::Single);
        border->set_LineWidth(30);
        border->set_Color(System::Drawing::Color::get_Blue());
        border->set_DistanceFromText(0);

        doc->Save(ArtifactsDir + u"PageSetup.PageBorderProperties.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.PageBorderProperties.docx");
        pageSetup = doc->get_FirstSection()->get_PageSetup();

        ASSERT_FALSE(pageSetup->get_BorderAlwaysInFront());
        ASSERT_EQ(PageBorderDistanceFrom::PageEdge, pageSetup->get_BorderDistanceFrom());
        ASSERT_EQ(PageBorderAppliesTo::FirstPage, pageSetup->get_BorderAppliesTo());

        border = pageSetup->get_Borders()->idx_get(BorderType::Top);

        ASSERT_EQ(LineStyle::Single, border->get_LineStyle());
        ASPOSE_ASSERT_EQ(30.0, border->get_LineWidth());
        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), border->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(0.0, border->get_DistanceFromText());
    }

    void PageBorders()
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

        pageSetup->get_Borders()->set_LineStyle(LineStyle::DoubleWave);
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
            ASSERT_EQ(LineStyle::DoubleWave, border->get_LineStyle());
            ASPOSE_ASSERT_EQ(2.0, border->get_LineWidth());
            ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), border->get_Color().ToArgb());
            ASPOSE_ASSERT_EQ(24.0, border->get_DistanceFromText());
            ASSERT_TRUE(border->get_Shadow());
        }
    }

    void PageNumbering()
    {
        //ExStart
        //ExFor:PageSetup.RestartPageNumbering
        //ExFor:PageSetup.PageStartingNumber
        //ExFor:PageSetup.PageNumberStyle
        //ExFor:DocumentBuilder.InsertField(String, String)
        //ExSummary:Shows how to set up page numbering in a section.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Section 1, page 1.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Section 1, page 2.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Section 1, page 3.");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Writeln(u"Section 2, page 1.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Section 2, page 2.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Section 2, page 3.");

        // Move the document builder to the first section's primary header,
        // which every page in that section will display.
        builder->MoveToSection(0);
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);

        // Insert a PAGE field, which will display the number of the current page.
        builder->Write(u"Page ");
        builder->InsertField(u"PAGE", u"");

        // Configure the section to have the page count that PAGE fields display start from 5.
        // Also, configure all PAGE fields to display their page numbers using uppercase Roman numerals.
        SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
        pageSetup->set_RestartPageNumbering(true);
        pageSetup->set_PageStartingNumber(5);
        pageSetup->set_PageNumberStyle(NumberStyle::UppercaseRoman);

        // Create another primary header for the second section, with another PAGE field.
        builder->MoveToSection(1);
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
        builder->Write(u" - ");
        builder->InsertField(u"PAGE", u"");
        builder->Write(u" - ");

        // Configure the section to have the page count that PAGE fields display start from 10.
        // Also, configure all PAGE fields to display their page numbers using Arabic numbers.
        pageSetup = doc->get_Sections()->idx_get(1)->get_PageSetup();
        pageSetup->set_PageStartingNumber(10);
        pageSetup->set_RestartPageNumbering(true);
        pageSetup->set_PageNumberStyle(NumberStyle::Arabic);

        doc->Save(ArtifactsDir + u"PageSetup.PageNumbering.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.PageNumbering.docx");
        pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();

        ASSERT_TRUE(pageSetup->get_RestartPageNumbering());
        ASSERT_EQ(5, pageSetup->get_PageStartingNumber());
        ASSERT_EQ(NumberStyle::UppercaseRoman, pageSetup->get_PageNumberStyle());

        pageSetup = doc->get_Sections()->idx_get(1)->get_PageSetup();

        ASSERT_TRUE(pageSetup->get_RestartPageNumbering());
        ASSERT_EQ(10, pageSetup->get_PageStartingNumber());
        ASSERT_EQ(NumberStyle::Arabic, pageSetup->get_PageNumberStyle());
    }

    void FootnoteOptions_()
    {
        //ExStart
        //ExFor:PageSetup.EndnoteOptions
        //ExFor:PageSetup.FootnoteOptions
        //ExSummary:Shows how to configure options affecting footnotes/endnotes in a section.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Hello world!");
        builder->InsertFootnote(FootnoteType::Footnote, u"Footnote reference text.");

        // Configure all footnotes in the first section to restart the numbering from 1
        // at each new page and display themselves directly beneath the text on every page.
        SharedPtr<FootnoteOptions> footnoteOptions = doc->get_Sections()->idx_get(0)->get_PageSetup()->get_FootnoteOptions();
        footnoteOptions->set_Position(FootnotePosition::BeneathText);
        footnoteOptions->set_RestartRule(FootnoteNumberingRule::RestartPage);
        footnoteOptions->set_StartNumber(1);

        builder->Write(u" Hello again.");
        builder->InsertFootnote(FootnoteType::Footnote, u"Endnote reference text.");

        // Configure all endnotes in the first section to maintain a continuous count throughout the section,
        // starting from 1. Also, set them all to appear collected at the end of the document.
        SharedPtr<EndnoteOptions> endnoteOptions = doc->get_Sections()->idx_get(0)->get_PageSetup()->get_EndnoteOptions();
        endnoteOptions->set_Position(EndnotePosition::EndOfDocument);
        endnoteOptions->set_RestartRule(FootnoteNumberingRule::Continuous);
        endnoteOptions->set_StartNumber(1);

        doc->Save(ArtifactsDir + u"PageSetup.FootnoteOptions.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.FootnoteOptions.docx");
        footnoteOptions = doc->get_FirstSection()->get_PageSetup()->get_FootnoteOptions();

        ASSERT_EQ(FootnotePosition::BeneathText, footnoteOptions->get_Position());
        ASSERT_EQ(FootnoteNumberingRule::RestartPage, footnoteOptions->get_RestartRule());
        ASSERT_EQ(1, footnoteOptions->get_StartNumber());

        endnoteOptions = doc->get_FirstSection()->get_PageSetup()->get_EndnoteOptions();

        ASSERT_EQ(EndnotePosition::EndOfDocument, endnoteOptions->get_Position());
        ASSERT_EQ(FootnoteNumberingRule::Continuous, endnoteOptions->get_RestartRule());
        ASSERT_EQ(1, endnoteOptions->get_StartNumber());
    }

    void Bidi(bool reverseColumns)
    {
        //ExStart
        //ExFor:PageSetup.Bidi
        //ExSummary:Shows how to set the order of text columns in a section.
        auto doc = MakeObject<Document>();

        SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
        pageSetup->get_TextColumns()->SetCount(3);

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Column 1.");
        builder->InsertBreak(BreakType::ColumnBreak);
        builder->Write(u"Column 2.");
        builder->InsertBreak(BreakType::ColumnBreak);
        builder->Write(u"Column 3.");

        // Set the "Bidi" property to "true" to arrange the columns starting from the page's right side.
        // The order of the columns will match the direction of the right-to-left text.
        // Set the "Bidi" property to "false" to arrange the columns starting from the page's left side.
        // The order of the columns will match the direction of the left-to-right text.
        pageSetup->set_Bidi(reverseColumns);

        doc->Save(ArtifactsDir + u"PageSetup.Bidi.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.Bidi.docx");
        pageSetup = doc->get_FirstSection()->get_PageSetup();

        ASSERT_EQ(3, pageSetup->get_TextColumns()->get_Count());
        ASPOSE_ASSERT_EQ(reverseColumns, pageSetup->get_Bidi());
    }

    void PageBorder()
    {
        //ExStart
        //ExFor:PageSetup.BorderSurroundsFooter
        //ExFor:PageSetup.BorderSurroundsHeader
        //ExSummary:Shows how to apply a border to the page and header/footer.
        auto doc = MakeObject<Document>();

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world! This is the main body text.");
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Write(u"This is the header.");
        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->Write(u"This is the footer.");
        builder->MoveToDocumentEnd();

        // Insert a blue double-line border.
        SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
        pageSetup->get_Borders()->set_LineStyle(LineStyle::Double);
        pageSetup->get_Borders()->set_Color(System::Drawing::Color::get_Blue());

        // A section's PageSetup object has "BorderSurroundsHeader" and "BorderSurroundsFooter" flags that determine
        // whether a page border surrounds the main body text, also includes the header or footer, respectively.
        // Set the "BorderSurroundsHeader" flag to "true" to surround the header with our border,
        // and then set the "BorderSurroundsFooter" flag to leave the footer outside of the border.
        pageSetup->set_BorderSurroundsHeader(true);
        pageSetup->set_BorderSurroundsFooter(false);

        doc->Save(ArtifactsDir + u"PageSetup.PageBorder.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.PageBorder.docx");
        pageSetup = doc->get_FirstSection()->get_PageSetup();

        ASSERT_TRUE(pageSetup->get_BorderSurroundsHeader());
        ASSERT_FALSE(pageSetup->get_BorderSurroundsFooter());
    }

    void Gutter()
    {
        //ExStart
        //ExFor:PageSetup.Gutter
        //ExFor:PageSetup.RtlGutter
        //ExFor:PageSetup.MultiplePages
        //ExSummary:Shows how to set gutter margins.
        auto doc = MakeObject<Document>();

        // Insert text that spans several pages.
        auto builder = MakeObject<DocumentBuilder>(doc);
        for (int i = 0; i < 6; i++)
        {
            builder->Write(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") +
                           u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder->InsertBreak(BreakType::PageBreak);
        }

        // A gutter adds whitespaces to either the left or right page margin,
        // which makes up for the center folding of pages in a book encroaching on the page's layout.
        SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();

        // Determine how much space our pages have for text within the margins and then add an amount to pad a margin.
        ASSERT_NEAR(470.30, pageSetup->get_PageWidth() - pageSetup->get_LeftMargin() - pageSetup->get_RightMargin(), 0.01);

        pageSetup->set_Gutter(100.0);

        // Set the "RtlGutter" property to "true" to place the gutter in a more suitable position for right-to-left text.
        pageSetup->set_RtlGutter(true);

        // Set the "MultiplePages" property to "MultiplePagesType.MirrorMargins" to alternate
        // the left/right page side position of margins every page.
        pageSetup->set_MultiplePages(MultiplePagesType::MirrorMargins);

        doc->Save(ArtifactsDir + u"PageSetup.Gutter.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.Gutter.docx");
        pageSetup = doc->get_FirstSection()->get_PageSetup();

        ASPOSE_ASSERT_EQ(100.0, pageSetup->get_Gutter());
        ASSERT_TRUE(pageSetup->get_RtlGutter());
        ASSERT_EQ(MultiplePagesType::MirrorMargins, pageSetup->get_MultiplePages());
    }

    void Booklet()
    {
        //ExStart
        //ExFor:PageSetup.Gutter
        //ExFor:PageSetup.MultiplePages
        //ExFor:PageSetup.SheetsPerBooklet
        //ExSummary:Shows how to configure a document that can be printed as a book fold.
        auto doc = MakeObject<Document>();

        // Insert text that spans 16 pages.
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"My Booklet:");

        for (int i = 0; i < 15; i++)
        {
            builder->InsertBreak(BreakType::PageBreak);
            builder->Write(String::Format(u"Booklet face #{0}", i));
        }

        // Configure the first section's "PageSetup" property to print the document in the form of a book fold.
        // When we print this document on both sides, we can take the pages to stack them
        // and fold them all down the middle at once. The contents of the document will line up into a book fold.
        SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
        pageSetup->set_MultiplePages(MultiplePagesType::BookFoldPrinting);

        // We can only specify the number of sheets in multiples of 4.
        pageSetup->set_SheetsPerBooklet(4);

        doc->Save(ArtifactsDir + u"PageSetup.Booklet.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.Booklet.docx");
        pageSetup = doc->get_FirstSection()->get_PageSetup();

        ASSERT_EQ(MultiplePagesType::BookFoldPrinting, pageSetup->get_MultiplePages());
        ASSERT_EQ(4, pageSetup->get_SheetsPerBooklet());
    }

    void SetTextOrientation()
    {
        //ExStart
        //ExFor:PageSetup.TextOrientation
        //ExSummary:Shows how to set text orientation.
        auto doc = MakeObject<Document>();

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // Set the "TextOrientation" property to "TextOrientation.Upward" to rotate all the text 90 degrees
        // to the right so that all left-to-right text now goes top-to-bottom.
        SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
        pageSetup->set_TextOrientation(TextOrientation::Upward);

        doc->Save(ArtifactsDir + u"PageSetup.SetTextOrientation.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"PageSetup.SetTextOrientation.docx");
        pageSetup = doc->get_FirstSection()->get_PageSetup();

        ASSERT_EQ(TextOrientation::Upward, pageSetup->get_TextOrientation());
    }

    //ExStart
    //ExFor:PageSetup.SuppressEndnotes
    //ExFor:Body.ParentSection
    //ExSummary:Shows how to store endnotes at the end of each section, and modify their positions.
    void SuppressEndnotes()
    {
        auto doc = MakeObject<Document>();
        doc->RemoveAllChildren();

        // By default, a document compiles all endnotes at its end.
        ASSERT_EQ(EndnotePosition::EndOfDocument, doc->get_EndnoteOptions()->get_Position());

        // We use the "Position" property of the document's "EndnoteOptions" object
        // to collect endnotes at the end of each section instead.
        doc->get_EndnoteOptions()->set_Position(EndnotePosition::EndOfSection);

        InsertSectionWithEndnote(doc, u"Section 1", u"Endnote 1, will stay in section 1");
        InsertSectionWithEndnote(doc, u"Section 2", u"Endnote 2, will be pushed down to section 3");
        InsertSectionWithEndnote(doc, u"Section 3", u"Endnote 3, will stay in section 3");

        // While getting sections to display their respective endnotes, we can set the "SuppressEndnotes" flag
        // of a section's "PageSetup" object to "true" to revert to the default behavior and pass its endnotes
        // onto the next section.
        SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(1)->get_PageSetup();
        pageSetup->set_SuppressEndnotes(true);

        doc->Save(ArtifactsDir + u"PageSetup.SuppressEndnotes.docx");
        TestSuppressEndnotes(MakeObject<Document>(ArtifactsDir + u"PageSetup.SuppressEndnotes.docx"));
        //ExSkip
    }

    /// <summary>
    /// Append a section with text and an endnote to a document.
    /// </summary>
    static void InsertSectionWithEndnote(SharedPtr<Document> doc, String sectionBodyText, String endnoteText)
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
        builder->InsertFootnote(FootnoteType::Endnote, endnoteText);
    }
    //ExEnd

    static void TestSuppressEndnotes(SharedPtr<Document> doc)
    {
        SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(1)->get_PageSetup();

        ASSERT_TRUE(pageSetup->get_SuppressEndnotes());
    }
};

} // namespace ApiExamples
