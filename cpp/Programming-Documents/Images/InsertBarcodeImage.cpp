#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Nodes/Node.h>
#include <Model/Sections/HeaderFooterType.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/SectionStart.h>
#include <Model/Text/ControlChar.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/Text/TabAlignment.h>
#include <Model/Text/TabLeader.h>
#include <Model/Text/TabStop.h>
#include <Model/Text/TabStopCollection.h>

using namespace Aspose::Words;

namespace
{
    // ExStart:InsertBarcodeIntoFooter
    void InsertBarcodeIntoFooter(System::SharedPtr<DocumentBuilder> builder, System::SharedPtr<Section> section, int32_t pageId, HeaderFooterType footerType, System::String const &inputDataDir)
    {
        // Move to the footer type in the specific section.
        builder->MoveToSection(section->get_Document()->IndexOf(section));
        builder->MoveToHeaderFooter(footerType);

        // Insert the barcode, then move to the next line and insert the ID along with the page number.
        // Use pageId if you need to insert a different barcode on each page. 0 = First page, 1 = Second page etc.
        builder->InsertImage(System::Drawing::Image::FromFile(inputDataDir + u"Barcode1.png"));
        builder->Writeln();
        builder->Write(u"1234567890");
        builder->InsertField(u"PAGE");

        // Create a right aligned tab at the right margin.
        double tabPos = section->get_PageSetup()->get_PageWidth() - section->get_PageSetup()->get_RightMargin() - section->get_PageSetup()->get_LeftMargin();
        builder->get_CurrentParagraph()->get_ParagraphFormat()->get_TabStops()->Add(System::MakeObject<TabStop>(tabPos, TabAlignment::Right, TabLeader::None));

        // Move to the right hand side of the page and insert the page and page total.
        builder->Write(ControlChar::Tab());
        builder->InsertField(u"PAGE");
        builder->Write(u" of ");
        builder->InsertField(u"NUMPAGES");
    }
    // ExEnd:InsertBarcodeIntoFooter
}

void InsertBarcodeImage()
{
    std::cout << "InsertBarcodeImage example started." << std::endl;
    // ExStart:InsertBarcodeImage
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithImages();
    System::String outputDataDir = GetOutputDataDir_WorkingWithImages();
    // Create a blank documenet.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // The number of pages the document should have.
    int32_t numPages = 4;
    // The document starts with one section, insert the barcode into this existing section.
    InsertBarcodeIntoFooter(builder, doc->get_FirstSection(), 1, HeaderFooterType::FooterPrimary, inputDataDir);

    for (int32_t i = 1; i < numPages; i++)
    {
        // Clone the first section and add it into the end of the document.
        System::SharedPtr<Section> cloneSection = System::DynamicCast<Section>(System::StaticCast<Node>(doc->get_FirstSection())->Clone(false));
        cloneSection->get_PageSetup()->set_SectionStart(SectionStart::NewPage);
        doc->AppendChild(cloneSection);

        // Insert the barcode and other information into the footer of the section.
        InsertBarcodeIntoFooter(builder, cloneSection, i, HeaderFooterType::FooterPrimary, inputDataDir);
    }

    System::String outputPath = outputDataDir + u"InsertBarcodeImage.docx";
    // Save the document as a DOCX to disk. You can also save this directly to a stream.
    doc->Save(outputPath);
    // ExEnd:InsertBarcodeImage
    std::cout << "Barcode image on each page of document inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertBarcodeImage example finished." << std::endl << std::endl;
}