#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace
{
    // ExStart:ConvertImageToPdf
    void ConvertImageToPdf(System::String const &inputFileName, System::String const &outputFileName)
    {
        std::cout << "Converting " << inputFileName.ToUtf8String() << " to PDF ...." << std::endl;
        // Create Document and DocumentBuilder. 
        // The builder makes it simple to add content to the document.
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Read the image from file, ensure it is disposed.
        System::SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromFile(inputFileName);
        {
			// Insert a section break before each new page, in case of a multi-frame TIFF.
			builder->InsertBreak(BreakType::SectionBreakNewPage);

			// We want the size of the page to be the same as the size of the image.
			// Convert pixels to points to size the page to the actual image size.
			System::SharedPtr<PageSetup> ps = builder->get_PageSetup();
			ps->set_PageWidth(ConvertUtil::PixelToPoint(image->get_Width(), image->get_HorizontalResolution()));
			ps->set_PageHeight(ConvertUtil::PixelToPoint(image->get_Height(), image->get_VerticalResolution()));

			// Insert the image into the document and position it at the top left corner of the page.
			builder->InsertImage(image, RelativeHorizontalPosition::Page, 0, RelativeVerticalPosition::Page, 0, ps->get_PageWidth(), ps->get_PageHeight(), WrapType::None);
        }

        // Save the document to PDF.
        doc->Save(outputFileName);
    }
    // ExEnd:ConvertImageToPdf
}

void ImageToPdf()
{
    std::cout << "ImageToPdf example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();
    // ExStart:ImageToPdf
    ConvertImageToPdf(inputDataDir + u"Test.jpg", outputDataDir + u"ImageToPdf_Jpg.pdf");
    ConvertImageToPdf(inputDataDir + u"Test.png", outputDataDir + u"ImageToPdf_Png.pdf");
    ConvertImageToPdf(inputDataDir + u"Test.wmf", outputDataDir + u"ImageToPdf_Wmf.pdf");
    ConvertImageToPdf(inputDataDir + u"Test.tiff", outputDataDir + u"ImageToPdf_Tiff.pdf");
    ConvertImageToPdf(inputDataDir + u"Test.gif", outputDataDir + u"ImageToPdf_Gif.pdf");
    // ExEnd:ImageToPdf
    std::cout << "ImageToPdf example finished." << std::endl << std::endl;
}