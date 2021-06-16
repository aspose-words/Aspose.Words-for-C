#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/Orientation.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Replacing/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Replacing/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Replacing/ReplaceAction.h>
#include <Aspose.Words.Cpp/Replacing/ReplacingArgs.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/ExportHeadersFootersMode.h>
#include <Aspose.Words.Cpp/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/Story.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/Tables/TableCollection.h>
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

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Replacing;

namespace ApiExamples {

class ExHeaderFooter : public ApiExampleBase
{
public:
    void Create()
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
        //ExSummary:Shows how to create a header and a footer.
        auto doc = MakeObject<Document>();

        // Create a header and append a paragraph to it. The text in that paragraph
        // will appear at the top of every page of this section, above the main body text.
        auto header = MakeObject<HeaderFooter>(doc, HeaderFooterType::HeaderPrimary);
        doc->get_FirstSection()->get_HeadersFooters()->Add(header);

        SharedPtr<Paragraph> para = header->AppendParagraph(u"My header.");

        ASSERT_TRUE(header->get_IsHeader());
        ASSERT_TRUE(para->get_IsEndOfHeaderFooter());

        // Create a footer and append a paragraph to it. The text in that paragraph
        // will appear at the bottom of every page of this section, below the main body text.
        auto footer = MakeObject<HeaderFooter>(doc, HeaderFooterType::FooterPrimary);
        doc->get_FirstSection()->get_HeadersFooters()->Add(footer);

        para = footer->AppendParagraph(u"My footer.");

        ASSERT_FALSE(footer->get_IsHeader());
        ASSERT_TRUE(para->get_IsEndOfHeaderFooter());

        ASPOSE_ASSERT_EQ(footer, para->get_ParentStory());
        ASPOSE_ASSERT_EQ(footer->get_ParentSection(), para->get_ParentSection());
        ASPOSE_ASSERT_EQ(footer->get_ParentSection(), header->get_ParentSection());

        doc->Save(ArtifactsDir + u"HeaderFooter.Create.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"HeaderFooter.Create.docx");

        ASSERT_TRUE(doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderPrimary)->get_Range()->get_Text().Contains(u"My header."));
        ASSERT_TRUE(doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary)->get_Range()->get_Text().Contains(u"My footer."));
    }

    void Link()
    {
        //ExStart
        //ExFor:HeaderFooter.IsLinkedToPrevious
        //ExFor:HeaderFooterCollection.Item(System.Int32)
        //ExFor:HeaderFooterCollection.LinkToPrevious(Aspose.Words.HeaderFooterType,System.Boolean)
        //ExFor:HeaderFooterCollection.LinkToPrevious(System.Boolean)
        //ExFor:HeaderFooter.ParentSection
        //ExSummary:Shows how to link headers and footers between sections.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Section 1");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Write(u"Section 2");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Write(u"Section 3");

        // Move to the first section and create a header and a footer. By default,
        // the header and the footer will only appear on pages in the section that contains them.
        builder->MoveToSection(0);

        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Write(u"This is the header, which will be displayed in sections 1 and 2.");

        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->Write(u"This is the footer, which will be displayed in sections 1, 2 and 3.");

        // We can link a section's headers/footers to the previous section's headers/footers
        // to allow the linking section to display the linked section's headers/footers.
        doc->get_Sections()->idx_get(1)->get_HeadersFooters()->LinkToPrevious(true);

        // Each section will still have its own header/footer objects. When we link sections,
        // the linking section will display the linked section's header/footers while keeping its own.
        ASPOSE_ASSERT_NE(doc->get_Sections()->idx_get(0)->get_HeadersFooters()->idx_get(0), doc->get_Sections()->idx_get(1)->get_HeadersFooters()->idx_get(0));
        ASPOSE_ASSERT_NE(doc->get_Sections()->idx_get(0)->get_HeadersFooters()->idx_get(0)->get_ParentSection(),
                         doc->get_Sections()->idx_get(1)->get_HeadersFooters()->idx_get(0)->get_ParentSection());

        // Link the headers/footers of the third section to the headers/footers of the second section.
        // The second section already links to the first section's header/footers,
        // so linking to the second section will create a link chain.
        // The first, second, and now the third sections will all display the first section's headers.
        doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LinkToPrevious(true);

        // We can un-link a previous section's header/footers by passing "false" when calling the LinkToPrevious method.
        doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LinkToPrevious(false);

        // We can also select only a specific type of header/footer to link using this method.
        // The third section now will have the same footer as the second and first sections, but not the header.
        doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LinkToPrevious(HeaderFooterType::FooterPrimary, true);

        // The first section's header/footers cannot link themselves to anything because there is no previous section.
        ASSERT_EQ(2, doc->get_Sections()->idx_get(0)->get_HeadersFooters()->get_Count());
        ASSERT_EQ(2, doc->get_Sections()->idx_get(0)->get_HeadersFooters()->LINQ_Count(
                         [](SharedPtr<Node> hf) { return !(System::DynamicCast<HeaderFooter>(hf))->get_IsLinkedToPrevious(); }));

        // All the second section's header/footers are linked to the first section's headers/footers.
        ASSERT_EQ(6, doc->get_Sections()->idx_get(1)->get_HeadersFooters()->get_Count());
        ASSERT_EQ(6, doc->get_Sections()->idx_get(1)->get_HeadersFooters()->LINQ_Count(
                         [](SharedPtr<Node> hf) { return (System::DynamicCast<HeaderFooter>(hf))->get_IsLinkedToPrevious(); }));

        // In the third section, only the footer is linked to the first section's footer via the second section.
        ASSERT_EQ(6, doc->get_Sections()->idx_get(2)->get_HeadersFooters()->get_Count());
        ASSERT_EQ(5, doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LINQ_Count(
                         [](SharedPtr<Node> hf) { return !(System::DynamicCast<HeaderFooter>(hf))->get_IsLinkedToPrevious(); }));
        ASSERT_TRUE(doc->get_Sections()->idx_get(2)->get_HeadersFooters()->idx_get(3)->get_IsLinkedToPrevious());

        doc->Save(ArtifactsDir + u"HeaderFooter.Link.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"HeaderFooter.Link.docx");

        ASSERT_EQ(2, doc->get_Sections()->idx_get(0)->get_HeadersFooters()->get_Count());
        ASSERT_EQ(2, doc->get_Sections()->idx_get(0)->get_HeadersFooters()->LINQ_Count(
                         [](SharedPtr<Node> hf) { return !(System::DynamicCast<HeaderFooter>(hf))->get_IsLinkedToPrevious(); }));

        ASSERT_EQ(0, doc->get_Sections()->idx_get(1)->get_HeadersFooters()->get_Count());
        ASSERT_EQ(0, doc->get_Sections()->idx_get(1)->get_HeadersFooters()->LINQ_Count(
                         static_cast<System::Func<SharedPtr<Node>, bool>>(static_cast<std::function<bool(SharedPtr<Node> hf)>>(
                             [](SharedPtr<Node> hf) -> bool { return (System::DynamicCast<HeaderFooter>(hf))->get_IsLinkedToPrevious(); }))));

        ASSERT_EQ(5, doc->get_Sections()->idx_get(2)->get_HeadersFooters()->get_Count());
        ASSERT_EQ(5, doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LINQ_Count(
                         [](SharedPtr<Node> hf) { return !(System::DynamicCast<HeaderFooter>(hf))->get_IsLinkedToPrevious(); }));
    }

    void RemoveFooters()
    {
        //ExStart
        //ExFor:Section.HeadersFooters
        //ExFor:HeaderFooterCollection
        //ExFor:HeaderFooterCollection.Item(HeaderFooterType)
        //ExFor:HeaderFooter
        //ExSummary:Shows how to delete all footers from a document.
        auto doc = MakeObject<Document>(MyDir + u"Header and footer types.docx");

        // Iterate through each section and remove footers of every kind.
        for (const auto& section : System::IterateOver(doc->LINQ_OfType<SharedPtr<Section>>()))
        {
            // There are three kinds of footer and header types.
            // 1 -  The "First" header/footer, which only appears on the first page of a section.
            SharedPtr<HeaderFooter> footer = section->get_HeadersFooters()->idx_get(HeaderFooterType::FooterFirst);
            if (footer != nullptr)
            {
                footer->Remove();
            }

            // 2 -  The "Primary" header/footer, which appears on odd pages.
            footer = section->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary);
            if (footer != nullptr)
            {
                footer->Remove();
            }

            // 3 -  The "Even" header/footer, which appears on odd even pages.
            footer = section->get_HeadersFooters()->idx_get(HeaderFooterType::FooterEven);
            if (footer != nullptr)
            {
                footer->Remove();
            }

            ASSERT_EQ(0,
                      section->get_HeadersFooters()->LINQ_Count([](SharedPtr<Node> hf) { return !(System::DynamicCast<HeaderFooter>(hf))->get_IsHeader(); }));
        }

        doc->Save(ArtifactsDir + u"HeaderFooter.RemoveFooters.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"HeaderFooter.RemoveFooters.docx");

        ASSERT_EQ(1, doc->get_Sections()->get_Count());
        ASSERT_EQ(0, doc->get_FirstSection()->get_HeadersFooters()->LINQ_Count(
                         static_cast<System::Func<SharedPtr<Node>, bool>>(static_cast<std::function<bool(SharedPtr<Node> hf)>>(
                             [](SharedPtr<Node> hf) -> bool { return !(System::DynamicCast<HeaderFooter>(hf))->get_IsHeader(); }))));
        ASSERT_EQ(3, doc->get_FirstSection()->get_HeadersFooters()->LINQ_Count(
                         static_cast<System::Func<SharedPtr<Node>, bool>>(static_cast<std::function<bool(SharedPtr<Node> hf)>>(
                             [](SharedPtr<Node> hf) -> bool { return (System::DynamicCast<HeaderFooter>(hf))->get_IsHeader(); }))));
    }

    void ExportMode()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportHeadersFootersMode
        //ExFor:ExportHeadersFootersMode
        //ExSummary:Shows how to omit headers/footers when saving a document to HTML.
        auto doc = MakeObject<Document>(MyDir + u"Header and footer types.docx");

        // This document contains headers and footers. We can access them via the "HeadersFooters" collection.
        ASSERT_EQ(u"First header", doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderFirst)->GetText().Trim());

        // Formats such as .html do not split the document into pages, so headers/footers will not function the same way
        // they would when we open the document as a .docx using Microsoft Word.
        // If we convert a document with headers/footers to html, the conversion will assimilate the headers/footers into body text.
        // We can use a SaveOptions object to omit headers/footers while converting to html.
        auto saveOptions = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        saveOptions->set_ExportHeadersFootersMode(ExportHeadersFootersMode::None);

        doc->Save(ArtifactsDir + u"HeaderFooter.ExportMode.html", saveOptions);

        // Open our saved document and verify that it does not contain the header's text
        doc = MakeObject<Document>(ArtifactsDir + u"HeaderFooter.ExportMode.html");

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
        //ExSummary:Shows how to replace text in a document's footer.
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
    //ExFor:PageSetup.DifferentFirstPageHeaderFooter
    //ExSummary:Shows how to track the order in which a text replacement operation traverses nodes.
    void Order(bool differentFirstPageHeaderFooter)
    {
        auto doc = MakeObject<Document>(MyDir + u"Header and footer types.docx");

        SharedPtr<Section> firstPageSection = doc->get_FirstSection();

        auto logger = MakeObject<ExHeaderFooter::ReplaceLog>();
        auto options = MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(logger);

        // Using a different header/footer for the first page will affect the search order.
        firstPageSection->get_PageSetup()->set_DifferentFirstPageHeaderFooter(differentFirstPageHeaderFooter);
        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"(header|footer)"), u"", options);

        if (differentFirstPageHeaderFooter)
        {
            ASSERT_EQ(u"First header\nFirst footer\nSecond header\nSecond footer\nThird header\nThird footer\n", logger->get_Text().Replace(u"\r", u""));
        }
        else
        {
            ASSERT_EQ(u"Third header\nFirst header\nThird footer\nFirst footer\nSecond header\nSecond footer\n", logger->get_Text().Replace(u"\r", u""));
        }
    }

    /// <summary>
    /// During a find-and-replace operation, records the contents of every node that has text that the operation 'finds',
    /// in the state it is in before the replacement takes place.
    /// This will display the order in which the text replacement operation traverses nodes.
    /// </summary>
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

        ReplaceLog() : mTextBuilder(MakeObject<System::Text::StringBuilder>())
        {
        }

    private:
        SharedPtr<System::Text::StringBuilder> mTextBuilder;
    };
    //ExEnd

    void Primer()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Section> currentSection = builder->get_CurrentSection();
        SharedPtr<PageSetup> pageSetup = currentSection->get_PageSetup();

        // Specify if we want headers/footers of the first page to be different from other pages.
        // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
        // different headers/footers for odd and even pages.
        pageSetup->set_DifferentFirstPageHeaderFooter(true);

        // Create header for the first page.
        pageSetup->set_HeaderDistance(20);
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderFirst);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);

        builder->get_Font()->set_Name(u"Arial");
        builder->get_Font()->set_Bold(true);
        builder->get_Font()->set_Size(14);
        builder->Write(u"Aspose.Words Header/Footer Creation Primer - Title Page.");

        // Create header for pages other than first.
        pageSetup->set_HeaderDistance(20);
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);

        // Insert an absolutely positioned image into the top/left corner of the header.
        // Distance from the top/left edges of the page is set to 10 points.
        String imageFileName = ImageDir + u"Logo.jpg";
        builder->InsertImage(imageFileName, RelativeHorizontalPosition::Page, 10.0, RelativeVerticalPosition::Page, 10.0, 50.0, 50.0, WrapType::Through);

        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);
        builder->Write(u"Aspose.Words Header/Footer Creation Primer.");

        // Create footer for pages other than first.
        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);

        // We use a table with two cells to make one part of the text on the line (with page numbering)
        // to be aligned left, and the other part of the text (with copyright) to be aligned right.
        builder->StartTable();

        builder->get_CellFormat()->ClearFormatting();

        builder->InsertCell();

        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100.0F / 3));

        // Insert page numbering text here.
        // It uses PAGE and NUMPAGES fields to auto calculate the current page number and a total number of pages.
        builder->Write(u"Page ");
        builder->InsertField(u"PAGE", u"");
        builder->Write(u" of ");
        builder->InsertField(u"NUMPAGES", u"");

        builder->get_CurrentParagraph()->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Left);

        builder->InsertCell();
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100.0F * 2 / 3));

        builder->Write(u"(C) 2001 Aspose Pty Ltd. All rights reserved.");

        builder->get_CurrentParagraph()->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);

        builder->EndRow();
        builder->EndTable();

        builder->MoveToDocumentEnd();
        builder->InsertBreak(BreakType::PageBreak);

        // Make section break to create a third page with a different page orientation.
        builder->InsertBreak(BreakType::SectionBreakNewPage);

        currentSection = builder->get_CurrentSection();
        pageSetup = currentSection->get_PageSetup();

        pageSetup->set_Orientation(Orientation::Landscape);

        // This section does not need different first page header/footer.
        // We need only one title page in the document and the header/footer for this page
        // has already been defined in the previous section.
        pageSetup->set_DifferentFirstPageHeaderFooter(false);

        // This section displays headers/footers from the previous section by default.
        // Call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this.
        // Page width is different for the new section and therefore we need to set
        // a different cell widths for a footer table.
        currentSection->get_HeadersFooters()->LinkToPrevious(false);

        // If we want to use the already existing header/footer set for this section
        // but with some minor modifications then it may be expedient to copy headers/footers
        // from the previous section and apply the necessary modifications where we want them.
        CopyHeadersFootersFromPreviousSection(currentSection);

        SharedPtr<HeaderFooter> primaryFooter = currentSection->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary);

        SharedPtr<Row> row = primaryFooter->get_Tables()->idx_get(0)->get_FirstRow();
        row->get_FirstCell()->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100.0F / 3));
        row->get_LastCell()->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100.0F * 2 / 3));

        doc->Save(ArtifactsDir + u"HeaderFooter.Primer.docx");
    }

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

        for (const auto& headerFooter : System::IterateOver(previousSection->get_HeadersFooters()->LINQ_OfType<SharedPtr<HeaderFooter>>()))
        {
            section->get_HeadersFooters()->Add(headerFooter->Clone(true));
        }
    }
};

} // namespace ApiExamples
