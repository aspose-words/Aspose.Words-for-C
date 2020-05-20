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
        // Clearing resources under 'using' statement
        System::Details::DisposeGuard<1> imageDisposeGuard({image});

        try
        {
            // Find which dimension the frames in this image represent. For example 
            // The frames of a BMP or TIFF are "page dimension" whereas frames of a GIF image are "time dimension".
            System::SharedPtr<System::Drawing::Imaging::FrameDimension> dimension = System::MakeObject<System::Drawing::Imaging::FrameDimension>(image->get_FrameDimensionsList()[0]);

            // Get the number of frames in the image.
            int32_t framesCount = image->GetFrameCount(dimension);

            // Loop through all frames.
            for (int32_t frameIdx = 0; frameIdx < framesCount; frameIdx++)
            {
                // Insert a section break before each new page, in case of a multi-frame TIFF.
                if (frameIdx != 0)
                {
                    builder->InsertBreak(BreakType::SectionBreakNewPage);
                }

                // Select active frame.
                image->SelectActiveFrame(dimension, frameIdx);

                // We want the size of the page to be the same as the size of the image.
                // Convert pixels to points to size the page to the actual image size.
                System::SharedPtr<PageSetup> ps = builder->get_PageSetup();
                ps->set_PageWidth(ConvertUtil::PixelToPoint(image->get_Width(), image->get_HorizontalResolution()));
                ps->set_PageHeight(ConvertUtil::PixelToPoint(image->get_Height(), image->get_VerticalResolution()));

                // Insert the image into the document and position it at the top left corner of the page.
                builder->InsertImage(image, RelativeHorizontalPosition::Page, 0, RelativeVerticalPosition::Page, 0, ps->get_PageWidth(), ps->get_PageHeight(), WrapType::None);
            }
        }
        catch (...)
        {
            imageDisposeGuard.SetCurrentException(std::current_exception());
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