#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Saving/ExportHeadersFootersMode.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/Orientation.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Story.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <system/collections/ienumerable.h>
#include <system/date_time.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/string_builder.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Replacing;

namespace ApiExamples {

class ExHeaderFooter : public ApiExampleBase
{
public:
    void HeaderFooterCreate()
    {
        //ExStart
        //ExFor:HeaderFooter
        //ExFor:HeaderFooter.#ctor(DocumentBase, HeaderFooterType)
        //ExFor:HeaderFooter.HeaderFooterType
        //ExFor:HeaderFooter.IsHeader
        //ExFor:HeaderFooterCollection
        //ExFor:Paragraph.IsEndOfHeaderFooter
        //ExFor:Paragraph.ParentSection
        //ExFor:Paragraph.ParentStory
        //ExFor:Story.AppendParagraph
        //ExSummary:Creates a header and footer using the document object model and insert them into a section.
        auto doc = MakeObject<Document>();

        auto header = MakeObject<HeaderFooter>(doc, HeaderFooterType::HeaderPrimary);
        doc->get_FirstSection()->get_HeadersFooters()->Add(header);

        // Add a paragraph with text to the footer
        SharedPtr<Paragraph> para = header->AppendParagraph(u"My header");

        ASSERT_TRUE(header->get_IsHeader());
        ASSERT_TRUE(para->get_IsEndOfHeaderFooter());

        auto footer = MakeObject<HeaderFooter>(doc, HeaderFooterType::FooterPrimary);
        doc->get_FirstSection()->get_HeadersFooters()->Add(footer);

        // Add a paragraph with text to the footer
        para = footer->AppendParagraph(u"My footer");

        ASSERT_FALSE(footer->get_IsHeader());
        ASSERT_TRUE(para->get_IsEndOfHeaderFooter());

        ASPOSE_ASSERT_EQ(footer, para->get_ParentStory());
        ASPOSE_ASSERT_EQ(footer->get_ParentSection(), para->get_ParentSection());
        ASPOSE_ASSERT_EQ(footer->get_ParentSection(), header->get_ParentSection());

        doc->Save(ArtifactsDir + u"HeaderFooter.HeaderFooterCreate.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"HeaderFooter.HeaderFooterCreate.docx");

        ASSERT_TRUE(doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderPrimary)->get_Range()->get_Text().Contains(u"My header"));
        ASSERT_TRUE(doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary)->get_Range()->get_Text().Contains(u"My footer"));
    }

    void HeaderFooterLink()
    {
        //ExStart
        //ExFor:HeaderFooter.IsLinkedToPrevious
        //ExFor:HeaderFooterCollection.Item(System.Int32)
        //ExFor:HeaderFooterCollection.LinkToPrevious(Aspose.Words.HeaderFooterType,System.Boolean)
        //ExFor:HeaderFooterCollection.LinkToPrevious(System.Boolean)
        //ExFor:HeaderFooter.ParentSection
        //ExSummary:Shows how to link header/footers between sections.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create three sections
        builder->Write(u"Section 1");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Write(u"Section 2");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Write(u"Section 3");

        // Create a header and footer in the first section and give them text
        builder->MoveToSection(0);

        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Write(u"This is the header, which will be displayed in sections 1 and 2.");

        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->Write(u"This is the footer, which will be displayed in sections 1, 2 and 3.");

        // If headers/footers are linked by the next section, they appear in that section also
        // The second section will display the header/footers of the first
        doc->get_Sections()->idx_get(1)->get_HeadersFooters()->LinkToPrevious(true);

        // However, the underlying headers/footers in the respective header/footer collections of the sections still remain different
        // Linking just overrides the existing headers/footers from the latter section
        ASSERT_EQ(doc->get_Sections()->idx_get(0)->get_HeadersFooters()->idx_get(0)->get_HeaderFooterType(),
                  doc->get_Sections()->idx_get(1)->get_HeadersFooters()->idx_get(0)->get_HeaderFooterType());
        ASPOSE_ASSERT_NE(doc->get_Sections()->idx_get(0)->get_HeadersFooters()->idx_get(0)->get_ParentSection(),
                         doc->get_Sections()->idx_get(1)->get_HeadersFooters()->idx_get(0)->get_ParentSection());
        ASSERT_NE(doc->get_Sections()->idx_get(0)->get_HeadersFooters()->idx_get(0)->GetText(),
                  doc->get_Sections()->idx_get(1)->get_HeadersFooters()->idx_get(0)->GetText());

        // Likewise, unlinking headers/footers makes them not appear
        doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LinkToPrevious(false);

        // We can also choose only certain header/footer types to get linked, like the footer in this case
        // The 3rd section now won't have the same header but will have the same footer as the 2nd and 1st sections
        doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LinkToPrevious(HeaderFooterType::FooterPrimary, true);

        // The first section's header/footers can't link themselves to anything because there is no previous section
        ASSERT_EQ(2, doc->get_Sections()->idx_get(0)->get_HeadersFooters()->get_Count());

        ASSERT_EQ(2, doc->get_Sections()->idx_get(0)->get_HeadersFooters()->LINQ_Count([](SharedPtr<Node> hf) { return !(System::DynamicCast<HeaderFooter>(hf))->get_IsLinkedToPrevious(); }));

        // All of the second section's header/footers are linked to those of the first
        ASSERT_EQ(6, doc->get_Sections()->idx_get(1)->get_HeadersFooters()->get_Count());

        ASSERT_EQ(6, doc->get_Sections()->idx_get(1)->get_HeadersFooters()->LINQ_Count([](SharedPtr<Node> hf) { return (System::DynamicCast<HeaderFooter>(hf))->get_IsLinkedToPrevious(); }));

        // In the third section, only the footer we explicitly linked is linked to that of the second, and consequently the first section
        ASSERT_EQ(6, doc->get_Sections()->idx_get(2)->get_HeadersFooters()->get_Count());

        ASSERT_EQ(5, doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LINQ_Count([](SharedPtr<Node> hf) { return !(System::DynamicCast<HeaderFooter>(hf))->get_IsLinkedToPrevious(); }));
        ASSERT_TRUE(doc->get_Sections()->idx_get(2)->get_HeadersFooters()->idx_get(3)->get_IsLinkedToPrevious());

        doc->Save(ArtifactsDir + u"HeaderFooter.HeaderFooterLink.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"HeaderFooter.HeaderFooterLink.docx");

        ASSERT_EQ(2, doc->get_Sections()->idx_get(0)->get_HeadersFooters()->get_Count());

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Node> hf)> _local_func_3 = [](SharedPtr<Node> hf)
        {
            return !(System::DynamicCast<HeaderFooter>(hf))->get_IsLinkedToPrevious();
        };

        ASSERT_EQ(2,
                  doc->get_Sections()->idx_get(0)->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(_local_func_3)));

        ASSERT_EQ(0, doc->get_Sections()->idx_get(1)->get_HeadersFooters()->get_Count());

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Node> hf)> _local_func_4 = [](SharedPtr<Node> hf)
        {
            return (System::DynamicCast<HeaderFooter>(hf))->get_IsLinkedToPrevious();
        };

        ASSERT_EQ(0,
                  doc->get_Sections()->idx_get(1)->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(_local_func_4)));

        ASSERT_EQ(5, doc->get_Sections()->idx_get(2)->get_HeadersFooters()->get_Count());

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Node> hf)> _local_func_5 = [](SharedPtr<Node> hf)
        {
            return !(System::DynamicCast<HeaderFooter>(hf))->get_IsLinkedToPrevious();
        };

        ASSERT_EQ(5,
                  doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(_local_func_5)));
    }

    void RemoveFooters()
    {
        //ExStart
        //ExFor:Section.HeadersFooters
        //ExFor:HeaderFooterCollection
        //ExFor:HeaderFooterCollection.Item(HeaderFooterType)
        //ExFor:HeaderFooter
        //ExSummary:Deletes all footers from all sections, but leaves headers intact.
        auto doc = MakeObject<Document>(MyDir + u"Header and footer types.docx");

        for (auto section : System::IterateOver(doc->LINQ_OfType<SharedPtr<Section>>()))
        {
            // Up to three different footers are possible in a section (for first, even and odd pages)
            // We check and delete all of them
            SharedPtr<HeaderFooter> footer = section->get_HeadersFooters()->idx_get(HeaderFooterType::FooterFirst);
            if (footer != nullptr)
            {
                footer->Remove();
            }

            // Primary footer is the footer used for odd pages
            footer = section->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary);
            if (footer != nullptr)
            {
                footer->Remove();
            }

            footer = section->get_HeadersFooters()->idx_get(HeaderFooterType::FooterEven);
            if (footer != nullptr)
            {
                footer->Remove();
            }

            // All footers have been removed from the section's HeaderFooter collection,
            // so every remaining node is a header and has the "IsHeader" flag set to true

            ASSERT_EQ(0, section->get_HeadersFooters()->LINQ_Count([](SharedPtr<Node> hf) { return !(System::DynamicCast<HeaderFooter>(hf))->get_IsHeader(); }));
        }

        doc->Save(ArtifactsDir + u"HeaderFooter.RemoveFooters.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"HeaderFooter.RemoveFooters.docx");

        ASSERT_EQ(1, doc->get_Sections()->get_Count());

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Node> hf)> _local_func_7 = [](SharedPtr<Node> hf)
        {
            return !(System::DynamicCast<HeaderFooter>(hf))->get_IsHeader();
        };

        ASSERT_EQ(0, doc->get_FirstSection()->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(_local_func_7)));

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Node> hf)> _local_func_8 = [](SharedPtr<Node> hf)
        {
            return (System::DynamicCast<HeaderFooter>(hf))->get_IsHeader();
        };

        ASSERT_EQ(3, doc->get_FirstSection()->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(_local_func_8)));
    }

    void SetExportHeadersFootersMode()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportHeadersFootersMode
        //ExFor:ExportHeadersFootersMode
        //ExSummary:Demonstrates how to disable the export of headers and footers when saving to HTML based formats.
        auto doc = MakeObject<Document>(MyDir + u"Header and footer types.docx");

        // This document contains headers and footers, whose text contents can be looked up like this
        ASSERT_EQ(u"First header", doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderFirst)->GetText().Trim());

        // Formats such as html do not have a pre-defined equivalent for Microsoft Word headers/footers
        // If we convert a document with headers and/or footers to html, they will be assimilated into body text
        // We can use a SaveOptions object to omit headers/footers while converting to html
        auto saveOptions = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        saveOptions->set_ExportHeadersFootersMode(ExportHeadersFootersMode::None);

        doc->Save(ArtifactsDir + u"HeaderFooter.DisableHeadersFooters.html", saveOptions);

        // Open our saved document and verify that it does not contain the header's text
        doc = MakeObject<Document>(ArtifactsDir + u"HeaderFooter.DisableHeadersFooters.html");

        ASSERT_FALSE(doc->get_Range()->get_Text().Contains(u"First header"));
        //ExEnd
    }

    void ReplaceText()
    {
        //ExStart
        //ExFor:Document.FirstSection
        //ExFor:Section.HeadersFooters
        //ExFor:HeaderFooterCollection.Item(HeaderFooterType)
        //ExFor:HeaderFooter
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExSummary:Shows how to replace text in the document footer.
        // Open the template document, containing obsolete copyright information in the footer
        auto doc = MakeObject<Document>(MyDir + u"Footer.docx");

        SharedPtr<HeaderFooterCollection> headersFooters = doc->get_FirstSection()->get_HeadersFooters();
        SharedPtr<HeaderFooter> footer = headersFooters->idx_get(HeaderFooterType::FooterPrimary);

        auto options = MakeObject<FindReplaceOptions>();
        options->set_MatchCase(false);
        options->set_FindWholeWordsOnly(false);

        int currentYear = System::DateTime::get_Now().get_Year();
        footer->get_Range()->Replace(u"(C) 2006 Aspose Pty Ltd.", String::Format(u"Copyright (C) {0} by Aspose Pty Ltd.", currentYear), options);

        doc->Save(ArtifactsDir + u"HeaderFooter.ReplaceText.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"HeaderFooter.ReplaceText.docx");

        ASSERT_TRUE(doc->get_Range()->get_Text().Contains(String::Format(u"Copyright (C) {0} by Aspose Pty Ltd.", currentYear)));
    }

    //ExStart
    //ExFor:IReplacingCallback
    //ExSummary:Show changes for headers and footers order.
    void HeaderFooterOrder()
    {
        auto doc = MakeObject<Document>(MyDir + u"Header and footer types.docx");

        // Assert that we use special header and footer for the first page
        // The order for this: first header\footer, even header\footer, primary header\footer
        SharedPtr<Section> firstPageSection = doc->get_FirstSection();
        ASPOSE_ASSERT_EQ(true, firstPageSection->get_PageSetup()->get_DifferentFirstPageHeaderFooter());

        auto logger = MakeObject<ExHeaderFooter::ReplaceLog>();
        auto options = MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(logger);
        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"(header|footer)"), u"", options);

        doc->Save(ArtifactsDir + u"HeaderFooter.HeaderFooterOrder.docx");

        ASSERT_EQ(String(u"First header\nFirst footer\nSecond header\nSecond footer\nThird header\n") + u"Third footer\n",
                  logger->get_Text().Replace(u"\r", u""));

        // Prepare our string builder for assert results without "DifferentFirstPageHeaderFooter"
        logger->ClearText();

        // Remove special first page
        // The order for this: primary header, default header, primary footer, default footer, even header\footer
        firstPageSection->get_PageSetup()->set_DifferentFirstPageHeaderFooter(false);
        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"(header|footer)"), u"", options);

        ASSERT_EQ(String(u"Third header\nFirst header\nThird footer\nFirst footer\nSecond header\n") + u"Second footer\n",
                  logger->get_Text().Replace(u"\r", u""));
    }

private:
    class ReplaceLog : public IReplacingCallback
    {
    public:
        String get_Text()
        {
            return mTextBuilder->ToString();
        }

        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            mTextBuilder->AppendLine(args->get_MatchNode()->GetText());
            return ReplaceAction::Skip;
        }

        void ClearText()
        {
            mTextBuilder->Clear();
        }

        ReplaceLog() : mTextBuilder(MakeObject<System::Text::StringBuilder>())
        {
        }

    private:
        SharedPtr<System::Text::StringBuilder> mTextBuilder;
    };
    //ExEnd

public:
    void Primer()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Section> currentSection = builder->get_CurrentSection();
        SharedPtr<PageSetup> pageSetup = currentSection->get_PageSetup();

        // Specify if we want headers/footers of the first page to be different from other pages
        // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
        // different headers/footers for odd and even pages
        pageSetup->set_DifferentFirstPageHeaderFooter(true);

        // --- Create header for the first page ---
        pageSetup->set_HeaderDistance(20);
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderFirst);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);

        // Set font properties for header text
        builder->get_Font()->set_Name(u"Arial");
        builder->get_Font()->set_Bold(true);
        builder->get_Font()->set_Size(14);
        // Specify header title for the first page
        builder->Write(u"Aspose.Words Header/Footer Creation Primer - Title Page.");

        // --- Create header for pages other than first ---
        pageSetup->set_HeaderDistance(20);
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);

        // Insert absolutely positioned image into the top/left corner of the header
        // Distance from the top/left edges of the page is set to 10 points
        String imageFileName = ImageDir + u"Logo.jpg";
        builder->InsertImage(imageFileName, RelativeHorizontalPosition::Page, 10.0, RelativeVerticalPosition::Page, 10.0, 50.0, 50.0, WrapType::Through);

        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);
        // Specify another header title for other pages
        builder->Write(u"Aspose.Words Header/Footer Creation Primer.");

        // --- Create footer for pages other than first ---
        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);

        // We use table with two cells to make one part of the text on the line (with page numbering)
        // to be aligned left, and the other part of the text (with copyright) to be aligned right
        builder->StartTable();

        // Clear table borders
        builder->get_CellFormat()->ClearFormatting();

        builder->InsertCell();

        // Set first cell to 1/3 of the page width
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100.0F / 3));

        // Insert page numbering text here
        // It uses PAGE and NUMPAGES fields to auto calculate current page number and total number of pages
        builder->Write(u"Page ");
        builder->InsertField(u"PAGE", u"");
        builder->Write(u" of ");
        builder->InsertField(u"NUMPAGES", u"");

        // Align this text to the left
        builder->get_CurrentParagraph()->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Left);

        builder->InsertCell();
        // Set the second cell to 2/3 of the page width
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100.0F * 2 / 3));

        builder->Write(u"(C) 2001 Aspose Pty Ltd. All rights reserved.");

        // Align this text to the right
        builder->get_CurrentParagraph()->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);

        builder->EndRow();
        builder->EndTable();

        builder->MoveToDocumentEnd();
        // Make page break to create a second page on which the primary headers/footers will be seen
        builder->InsertBreak(BreakType::PageBreak);

        // Make section break to create a third page with different page orientation
        builder->InsertBreak(BreakType::SectionBreakNewPage);

        // Get the new section and its page setup
        currentSection = builder->get_CurrentSection();
        pageSetup = currentSection->get_PageSetup();

        // Set page orientation of the new section to landscape
        pageSetup->set_Orientation(Orientation::Landscape);

        // This section does not need different first page header/footer
        // We need only one title page in the document and the header/footer for this page
        // has already been defined in the previous section
        pageSetup->set_DifferentFirstPageHeaderFooter(false);

        // This section displays headers/footers from the previous section by default
        // Call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this
        // Page width is different for the new section and therefore we need to set
        // a different cell widths for a footer table
        currentSection->get_HeadersFooters()->LinkToPrevious(false);

        // If we want to use the already existing header/footer set for this section
        // but with some minor modifications then it may be expedient to copy headers/footers
        // from the previous section and apply the necessary modifications where we want them
        CopyHeadersFootersFromPreviousSection(currentSection);

        // Find the footer that we want to change
        SharedPtr<HeaderFooter> primaryFooter = currentSection->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary);

        SharedPtr<Row> row = primaryFooter->get_Tables()->idx_get(0)->get_FirstRow();
        row->get_FirstCell()->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100.0F / 3));
        row->get_LastCell()->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100.0F * 2 / 3));

        // Save the resulting document
        doc->Save(ArtifactsDir + u"HeaderFooter.Primer.docx");
    }

protected:
    /// <summary>
    /// Clones and copies headers/footers form the previous section to the specified section.
    /// </summary>
    static void CopyHeadersFootersFromPreviousSection(SharedPtr<Section> section)
    {
        auto previousSection = System::DynamicCast<Section>(section->get_PreviousSibling());

        if (previousSection == nullptr)
        {
            return;
        }

        section->get_HeadersFooters()->Clear();

        for (auto headerFooter : System::IterateOver(previousSection->get_HeadersFooters()->LINQ_OfType<SharedPtr<HeaderFooter>>()))
        {
            section->get_HeadersFooters()->Add(headerFooter->Clone(true));
        }
    }
};

} // namespace ApiExamples
