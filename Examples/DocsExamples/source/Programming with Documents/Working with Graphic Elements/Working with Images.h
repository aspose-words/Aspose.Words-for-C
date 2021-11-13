#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/ControlChar.h>
#include <Aspose.Words.Cpp/ConvertUtil.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/ImageSize.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Layout/LayoutCollector.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeList.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionStart.h>
#include <Aspose.Words.Cpp/TabAlignment.h>
#include <Aspose.Words.Cpp/TabLeader.h>
#include <Aspose.Words.Cpp/TabStop.h>
#include <Aspose.Words.Cpp/TabStopCollection.h>
#include <drawing/image.h>
#include <drawing/imaging/image_codec_info.h>
#include <drawing/imaging/image_format.h>
#include <drawing/size_f.h>
#include <system/collections/ienumerator.h>
#include <system/diagnostics/debug.h>
#include <system/exceptions.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Layout;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Graphic_Elements {

class Resampler : public System::Object
{
public:
    /// <summary>
    /// Resamples all images in the document that are greater than the specified PPI (pixels per inch) to the specified PPI
    /// and converts them to JPEG with the specified quality setting.
    /// </summary>
    /// <param name="doc">The document to process.</param>
    /// <param name="desiredPpi">Desired pixels per inch. 220 high quality. 150 screen quality. 96 email quality.</param>
    /// <param name="jpegQuality">0 - 100% JPEG quality.</param>
    /// <returns></returns>
    static int Resample(SharedPtr<Document> doc, int desiredPpi, int jpegQuality);

private:
    /// <summary>
    /// Resamples one VML or DrawingML image.
    /// </summary>
    static bool ResampleCore(SharedPtr<ImageData> imageData, System::Drawing::SizeF shapeSizeInPoints, int ppi, int jpegQuality);
    /// <summary>
    /// Gets the codec info for the specified image format.
    /// Throws if cannot find.
    /// </summary>
    static SharedPtr<System::Drawing::Imaging::ImageCodecInfo> GetEncoderInfo(SharedPtr<System::Drawing::Imaging::ImageFormat> format);
};

class WorkingWithImages : public DocsExamplesBase
{
public:
    void AddImageToEachPage()
    {
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Create and attach collector before the document before page layout is built.
        auto layoutCollector = MakeObject<LayoutCollector>(doc);

        // Images in a document are added to paragraphs to add an image to every page we need
        // to find at any paragraph belonging to each page.
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Node>>> enumerator = doc->SelectNodes(u"//Paragraph")->GetEnumerator();

        for (int page = 1; page <= doc->get_PageCount(); page++)
        {
            while (enumerator->MoveNext())
            {
                // Check if the current paragraph belongs to the target page.
                auto paragraph = System::DynamicCast<Paragraph>(enumerator->get_Current());
                if (layoutCollector->GetStartPageIndex(paragraph) == page)
                {
                    AddImageToPage(paragraph, page, ImagesDir);
                    break;
                }
            }
        }

        // If we need to save the document as a PDF or image, call UpdatePageLayout() method.
        doc->UpdatePageLayout();

        doc->Save(ArtifactsDir + u"WorkingWithImages.AddImageToEachPage.docx");
    }

    /// <summary>
    /// Adds an image to a page using the supplied paragraph.
    /// </summary>
    /// <param name="para">The paragraph to an an image to.</param>
    /// <param name="page">The page number the paragraph appears on.</param>
    void AddImageToPage(SharedPtr<Paragraph> para, int page, String imagesDir)
    {
        ASPOSE_UNUSED(imagesDir);
        auto doc = System::DynamicCast<Document>(para->get_Document());

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveTo(para);

        // Insert a logo to the top left of the page to place it in front of all other text.
        builder->InsertImage(ImagesDir + u"Transparent background logo.png", RelativeHorizontalPosition::Page, 60.0, RelativeVerticalPosition::Page, 60.0, -1.0,
                             -1.0, WrapType::None);

        // Insert a textbox next to the image which contains some text consisting of the page number.
        auto textBox = MakeObject<Shape>(doc, ShapeType::TextBox);

        // We want a floating shape relative to the page.
        textBox->set_WrapType(WrapType::None);
        textBox->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);
        textBox->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);

        textBox->set_Height(30);
        textBox->set_Width(200);
        textBox->set_Left(150);
        textBox->set_Top(80);

        textBox->AppendChild(MakeObject<Paragraph>(doc));
        builder->InsertNode(textBox);
        builder->MoveTo(textBox->get_FirstChild());
        builder->Writeln(String(u"This is a custom note for page ") + page);
    }

    void InsertBarcodeImage()
    {
        //ExStart:InsertBarcodeImage
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // The number of pages the document should have
        const int numPages = 4;
        // The document starts with one section, insert the barcode into this existing section
        InsertBarcodeIntoFooter(builder, doc->get_FirstSection(), HeaderFooterType::FooterPrimary);

        for (int i = 1; i < numPages; i++)
        {
            // Clone the first section and add it into the end of the document
            auto cloneSection = System::DynamicCast<Section>(doc->get_FirstSection()->Clone(false));
            cloneSection->get_PageSetup()->set_SectionStart(SectionStart::NewPage);
            doc->AppendChild(cloneSection);

            // Insert the barcode and other information into the footer of the section
            InsertBarcodeIntoFooter(builder, cloneSection, HeaderFooterType::FooterPrimary);
        }

        // Save the document as a PDF to disk
        // You can also save this directly to a stream
        doc->Save(ArtifactsDir + u"InsertBarcodeImage.docx");
        //ExEnd:InsertBarcodeImage
    }

    void CompressImages()
    {
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        // 220ppi Print - said to be excellent on most printers and screens.
        // 150ppi Screen - said to be good for web pages and projectors.
        // 96ppi Email - said to be good for minimal document size and sharing.
        const int desiredPpi = 150;

        // In .NET this seems to be a good compression/quality setting.
        const int jpegQuality = 90;

        // Resample images to the desired PPI and save.
        int count = Resampler::Resample(doc, desiredPpi, jpegQuality);

        std::cout << "Resampled " << count << " images." << std::endl;

        if (count != 1)
        {
            std::cout << "We expected to have only 1 image resampled in this test document!" << std::endl;
        }

        doc->Save(ArtifactsDir + u"CompressImages.docx");

        // Verify that the first image was compressed by checking the new PPI.
        doc = MakeObject<Document>(ArtifactsDir + u"CompressImages.docx");

        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        double imagePpi = shape->get_ImageData()->get_ImageSize()->get_WidthPixels() / ConvertUtil::PointToInch(shape->get_SizeInPoints().get_Width());

        CODEPORTING_DEBUG_ASSERT2(imagePpi < 150, u"Image was not resampled successfully.");
    }

protected:
    void InsertBarcodeIntoFooter(SharedPtr<DocumentBuilder> builder, SharedPtr<Section> section, HeaderFooterType footerType)
    {
        // Move to the footer type in the specific section.
        builder->MoveToSection(section->get_Document()->IndexOf(section));
        builder->MoveToHeaderFooter(footerType);

        // Insert the barcode, then move to the next line and insert the ID along with the page number.
        // Use pageId if you need to insert a different barcode on each page. 0 = First page, 1 = Second page etc.
        builder->InsertImage(System::Drawing::Image::FromFile(ImagesDir + u"Barcode.png"));
        builder->Writeln();
        builder->Write(u"1234567890");
        builder->InsertField(u"PAGE");

        // Create a right-aligned tab at the right margin.
        double tabPos = section->get_PageSetup()->get_PageWidth() - section->get_PageSetup()->get_RightMargin() - section->get_PageSetup()->get_LeftMargin();
        builder->get_CurrentParagraph()->get_ParagraphFormat()->get_TabStops()->Add(MakeObject<TabStop>(tabPos, TabAlignment::Right, TabLeader::None));

        // Move to the right-hand side of the page and insert the page and page total.
        builder->Write(ControlChar::Tab());
        builder->InsertField(u"PAGE");
        builder->Write(u" of ");
        builder->InsertField(u"NUMPAGES");
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements
