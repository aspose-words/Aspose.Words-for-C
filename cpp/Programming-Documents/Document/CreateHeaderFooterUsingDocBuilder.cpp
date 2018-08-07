#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/collections/ienumerator.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/Text/ParagraphAlignment.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/Font.h>
#include <Model/Tables/TableCollection.h>
#include <Model/Tables/Table.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/PreferredWidth.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Tables/Cell.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Sections/Orientation.h>
#include <Model/Sections/HeaderFooterType.h>
#include <Model/Sections/HeaderFooterCollection.h>
#include <Model/Sections/HeaderFooter.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/Node.h>
#include <Model/Fields/Field.h>
#include <Model/Drawing/WrapType.h>
#include <Model/Drawing/Shape.h>
#include <Model/Drawing/RelativeVerticalPosition.h>
#include <Model/Drawing/RelativeHorizontalPosition.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Document/BreakType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Drawing;

namespace
{
    // ExStart:CopyHeadersFootersFromPreviousSection
    /// <summary>
    /// Clones and copies headers/footers form the previous section to the specified section.
    /// </summary>
    void CopyHeadersFootersFromPreviousSection(const System::SharedPtr<Section>& section)
    {
        System::SharedPtr<Section> previousSection = System::DynamicCast<Aspose::Words::Section>(section->get_PreviousSibling());

        if (previousSection == nullptr)
        {
            return;
        }

        section->get_HeadersFooters()->Clear();

        auto headerFooter_enumerator = previousSection->get_HeadersFooters()->GetEnumerator();
        System::SharedPtr<HeaderFooter> headerFooter;
        while (headerFooter_enumerator->MoveNext() && (headerFooter = System::DynamicCast<HeaderFooter>(headerFooter_enumerator->get_Current()), true))
        {
            section->get_HeadersFooters()->Add((System::StaticCast<Aspose::Words::Node>(headerFooter))->Clone(true));
        }
    }
    // ExEnd:CopyHeadersFootersFromPreviousSection
}

void CreateHeaderFooterUsingDocBuilder()
{
    std::cout << "CreateHeaderFooterUsingDocBuilder example started." << std::endl;
    // ExStart:CreateHeaderFooterUsingDocBuilder
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();

    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    System::SharedPtr<Section> currentSection = builder->get_CurrentSection();
    System::SharedPtr<PageSetup> pageSetup = currentSection->get_PageSetup();

    // Specify if we want headers/footers of the first page to be different from other pages.
    // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
    // Different headers/footers for odd and even pages.
    pageSetup->set_DifferentFirstPageHeaderFooter(true);

    // --- Create header for the first page. ---
    pageSetup->set_HeaderDistance(20);
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderFirst);
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);

    // Set font properties for header text.
    builder->get_Font()->set_Name(u"Arial");
    builder->get_Font()->set_Bold(true);
    builder->get_Font()->set_Size(14);
    // Specify header title for the first page.
    builder->Write(u"Aspose.Words Header/Footer Creation Primer - Title Page.");

    // --- Create header for pages other than first. ---
    pageSetup->set_HeaderDistance(20);
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);

    // Insert absolutely positioned image into the top/left corner of the header.
    // Distance from the top/left edges of the page is set to 10 points.
    System::String imageFileName = dataDir + u"Aspose.Words.gif";
    builder->InsertImage(imageFileName, Aspose::Words::Drawing::RelativeHorizontalPosition::Page, 10, Aspose::Words::Drawing::RelativeVerticalPosition::Page, 10, 50, 50, Aspose::Words::Drawing::WrapType::Through);

    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Right);
    // Specify another header title for other pages.
    builder->Write(u"Aspose.Words Header/Footer Creation Primer.");

    // --- Create footer for pages other than first. ---
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterPrimary);

    // We use table with two cells to make one part of the text on the line (with page numbering)
    // To be aligned left, and the other part of the text (with copyright) to be aligned right.
    builder->StartTable();

    // Clear table borders.
    builder->get_CellFormat()->ClearFormatting();

    builder->InsertCell();

    // Set first cell to 1/3 of the page width.
    builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100 / 3));

    // Insert page numbering text here.
    // It uses PAGE and NUMPAGES fields to auto calculate current page number and total number of pages.
    builder->Write(u"Page ");
    builder->InsertField(u"PAGE", u"");
    builder->Write(u" of ");
    builder->InsertField(u"NUMPAGES", u"");

    // Align this text to the left.
    builder->get_CurrentParagraph()->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Left);

    builder->InsertCell();
    // Set the second cell to 2/3 of the page width.
    builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100 * 2 / 3));

    builder->Write(u"(C) 2001 Aspose Pty Ltd. All rights reserved.");

    // Align this text to the right.
    builder->get_CurrentParagraph()->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Right);

    builder->EndRow();
    builder->EndTable();

    builder->MoveToDocumentEnd();
    // Make page break to create a second page on which the primary headers/footers will be seen.
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);

    // Make section break to create a third page with different page orientation.
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);

    // Get the new section and its page setup.
    currentSection = builder->get_CurrentSection();
    pageSetup = currentSection->get_PageSetup();

    // Set page orientation of the new section to landscape.
    pageSetup->set_Orientation(Aspose::Words::Orientation::Landscape);

    // This section does not need different first page header/footer.
    // We need only one title page in the document and the header/footer for this page
    // Has already been defined in the previous section
    pageSetup->set_DifferentFirstPageHeaderFooter(false);

    // This section displays headers/footers from the previous section by default.
    // Call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this.
    // Page width is different for the new section and therefore we need to set 
    // A different cell widths for a footer table.
    currentSection->get_HeadersFooters()->LinkToPrevious(false);

    // If we want to use the already existing header/footer set for this section 
    // But with some minor modifications then it may be expedient to copy headers/footers
    // From the previous section and apply the necessary modifications where we want them.
    CopyHeadersFootersFromPreviousSection(currentSection);

    // Find the footer that we want to change.
    System::SharedPtr<HeaderFooter> primaryFooter = currentSection->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary);

    System::SharedPtr<Row> row = primaryFooter->get_Tables()->idx_get(0)->get_FirstRow();
    row->get_FirstCell()->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100 / 3));
    row->get_LastCell()->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100 * 2 / 3));

    System::String outputPath = dataDir + GetOutputFilePath(u"CreateHeaderFooterUsingDocBuilder.doc");
    // Save the resulting document.
    doc->Save(outputPath);
    // ExEnd:CreateHeaderFooterUsingDocBuilder
    std::cout << "Header and footer created successfully using document builder." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CreateHeaderFooterUsingDocBuilder example finished." << std::endl;
}
