#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Orientation.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Drawing;

namespace
{
    // ExStart:CopyHeadersFootersFromPreviousSection
    // Clones and copies headers/footers form the previous section to the specified section.
    void CopyHeadersFootersFromPreviousSection(const System::SharedPtr<Section>& section)
    {
        System::SharedPtr<Section> previousSection = System::DynamicCast<Section>(section->get_PreviousSibling());

        if (previousSection == nullptr)
        {
            return;
        }

        section->get_HeadersFooters()->Clear();

        for (System::SharedPtr<Node> headerFooterNode : System::IterateOver(previousSection->get_HeadersFooters()))
        {
            section->get_HeadersFooters()->Add(headerFooterNode->Clone(true));
        }
    }
    // ExEnd:CopyHeadersFootersFromPreviousSection
}

void CreateHeaderFooterUsingDocBuilder()
{
    std::cout << "CreateHeaderFooterUsingDocBuilder example started." << std::endl;
    // ExStart:CreateHeaderFooterUsingDocBuilder
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

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
    builder->MoveToHeaderFooter(HeaderFooterType::HeaderFirst);
    builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);

    // Set font properties for header text.
    builder->get_Font()->set_Name(u"Arial");
    builder->get_Font()->set_Bold(true);
    builder->get_Font()->set_Size(14);
    // Specify header title for the first page.
    builder->Write(u"Aspose.Words Header/Footer Creation Primer - Title Page.");

    // --- Create header for pages other than first. ---
    pageSetup->set_HeaderDistance(20);
    builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);

    // Insert absolutely positioned image into the top/left corner of the header.
    // Distance from the top/left edges of the page is set to 10 points.
    System::String imageFileName = inputDataDir + u"Aspose.Words.gif";
    builder->InsertImage(imageFileName, RelativeHorizontalPosition::Page, 10, RelativeVerticalPosition::Page, 10, 50, 50, WrapType::Through);

    builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);
    // Specify another header title for other pages.
    builder->Write(u"Aspose.Words Header/Footer Creation Primer.");

    // --- Create footer for pages other than first. ---
    builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);

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
    builder->get_CurrentParagraph()->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Left);

    builder->InsertCell();
    // Set the second cell to 2/3 of the page width.
    builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100 * 2 / 3));

    builder->Write(u"(C) 2001 Aspose Pty Ltd. All rights reserved.");

    // Align this text to the right.
    builder->get_CurrentParagraph()->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);

    builder->EndRow();
    builder->EndTable();

    builder->MoveToDocumentEnd();
    // Make page break to create a second page on which the primary headers/footers will be seen.
    builder->InsertBreak(BreakType::PageBreak);

    // Make section break to create a third page with different page orientation.
    builder->InsertBreak(BreakType::SectionBreakNewPage);

    // Get the new section and its page setup.
    currentSection = builder->get_CurrentSection();
    pageSetup = currentSection->get_PageSetup();

    // Set page orientation of the new section to landscape.
    pageSetup->set_Orientation(Orientation::Landscape);

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
    System::SharedPtr<HeaderFooter> primaryFooter = currentSection->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary);

    System::SharedPtr<Row> row = primaryFooter->get_Tables()->idx_get(0)->get_FirstRow();
    row->get_FirstCell()->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100 / 3));
    row->get_LastCell()->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(100 * 2 / 3));

    System::String outputPath = outputDataDir + u"CreateHeaderFooterUsingDocBuilder.doc";
    // Save the resulting document.
    doc->Save(outputPath);
    // ExEnd:CreateHeaderFooterUsingDocBuilder
    std::cout << "Header and footer created successfully using document builder." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CreateHeaderFooterUsingDocBuilder example finished." << std::endl;
}
