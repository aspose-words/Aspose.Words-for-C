#pragma once

#include <cstdint>
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
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/Tables/TableCollection.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithHeadersAndFooters : public DocsExamplesBase
{
public:
    void CreateHeaderFooter()
    {
        //ExStart:CreateHeaderFooterUsingDocBuilder
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Section> currentSection = builder->get_CurrentSection();
        SharedPtr<PageSetup> pageSetup = currentSection->get_PageSetup();
        // Specify if we want headers/footers of the first page to be different from other pages.
        // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
        // different headers/footers for odd and even pages.
        pageSetup->set_DifferentFirstPageHeaderFooter(true);
        pageSetup->set_HeaderDistance(20);

        builder->MoveToHeaderFooter(HeaderFooterType::HeaderFirst);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);

        builder->get_Font()->set_Name(u"Arial");
        builder->get_Font()->set_Bold(true);
        builder->get_Font()->set_Size(14);

        builder->Write(u"Aspose.Words Header/Footer Creation Primer - Title Page.");

        pageSetup->set_HeaderDistance(20);
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);

        // Insert a positioned image into the top/left corner of the header.
        // Distance from the top/left edges of the page is set to 10 points.
        builder->InsertImage(ImagesDir + u"Graphics Interchange Format.gif", RelativeHorizontalPosition::Page, 10.0, RelativeVerticalPosition::Page, 10.0, 50.0,
                             50.0, WrapType::Through);

        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);

        builder->Write(u"Aspose.Words Header/Footer Creation Primer.");

        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);

        // We use a table with two cells to make one part of the text on the line (with page numbering).
        // To be aligned left, and the other part of the text (with copyright) to be aligned right.
        builder->StartTable();

        builder->get_CellFormat()->ClearFormatting();

        builder->InsertCell();

        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100 / 3));

        // It uses PAGE and NUMPAGES fields to auto calculate the current page number and many pages.
        builder->Write(u"Page ");
        builder->InsertField(u"PAGE", u"");
        builder->Write(u" of ");
        builder->InsertField(u"NUMPAGES", u"");

        builder->get_CurrentParagraph()->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Left);

        builder->InsertCell();

        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100 * 2 / 3));

        builder->Write(u"(C) 2001 Aspose Pty Ltd. All rights reserved.");

        builder->get_CurrentParagraph()->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);

        builder->EndRow();
        builder->EndTable();

        builder->MoveToDocumentEnd();

        // Make a page break to create a second page on which the primary headers/footers will be seen.
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::SectionBreakNewPage);

        currentSection = builder->get_CurrentSection();
        pageSetup = currentSection->get_PageSetup();
        pageSetup->set_Orientation(Orientation::Landscape);
        // This section does not need a different first-page header/footer we need only one title page in the document,
        // and the header/footer for this page has already been defined in the previous section.
        pageSetup->set_DifferentFirstPageHeaderFooter(false);

        // This section displays headers/footers from the previous section
        // by default call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this page width
        // is different for the new section, and therefore we need to set different cell widths for a footer table.
        currentSection->get_HeadersFooters()->LinkToPrevious(false);

        // If we want to use the already existing header/footer set for this section.
        // But with some minor modifications, then it may be expedient to copy headers/footers
        // from the previous section and apply the necessary modifications where we want them.
        CopyHeadersFootersFromPreviousSection(currentSection);

        SharedPtr<HeaderFooter> primaryFooter = currentSection->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary);

        SharedPtr<Row> row = primaryFooter->get_Tables()->idx_get(0)->get_FirstRow();
        row->get_FirstCell()->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100 / 3));
        row->get_LastCell()->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100 * 2 / 3));

        doc->Save(ArtifactsDir + u"WorkingWithHeadersAndFooters.CreateHeaderFooter.docx");
        //ExEnd:CreateHeaderFooterUsingDocBuilder
    }

protected:
    /// <summary>
    /// Clones and copies headers/footers form the previous section to the specified section.
    /// </summary>
    void CopyHeadersFootersFromPreviousSection(SharedPtr<Section> section)
    {
        auto previousSection = System::DynamicCast<Section>(section->get_PreviousSibling());

        if (previousSection == nullptr)
        {
            return;
        }

        section->get_HeadersFooters()->Clear();

        for (const auto& headerFooter : System::IterateOver<HeaderFooter>(previousSection->get_HeadersFooters()))
        {
            section->get_HeadersFooters()->Add((System::StaticCast<Node>(headerFooter))->Clone(true));
        }
    }
};

}} // namespace DocsExamples::Programming_with_Documents
