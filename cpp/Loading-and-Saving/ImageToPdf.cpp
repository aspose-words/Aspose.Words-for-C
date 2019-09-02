#include "stdafx.h"
#include "examples.h"

#include <Model/Document/BreakType.h>
#include <Model/Document/ConvertUtil.h>
#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Drawing/RelativeHorizontalPosition.h>
#include <Model/Drawing/RelativeVerticalPosition.h>
#include <Model/Drawing/WrapType.h>
#include <Model/Sections/PageSetup.h>

/*using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace
{
    void ConvertImageToPdf(System::String const &inputFileName, System::String const &outputFileName)
    {
        std::cout << "Converting " << inputFileName.ToUtf8String() << " to PDF ...." << std::endl;
        // ExStart:ConvertImageToPdf
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
        // ExEnd:ConvertImageToPdf
    }
}

void ImageToPdf()
{
    std::cout << "ImageToPdf example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();
    ConvertImageToPdf(dataDir + u"Test.jpg", dataDir + GetOutputFilePath(u"ImageToPdf_Jpg.pdf"));
    ConvertImageToPdf(dataDir + u"Test.png", dataDir + GetOutputFilePath(u"ImageToPdf_Png.pdf"));
    ConvertImageToPdf(dataDir + u"Test.wmf", dataDir + GetOutputFilePath(u"ImageToPdf_Wmf.pdf"));
    ConvertImageToPdf(dataDir + u"Test.tiff", dataDir +GetOutputFilePath(u"ImageToPdf_Tiff.pdf"));
    ConvertImageToPdf(dataDir + u"Test.gif", dataDir + GetOutputFilePath(u"ImageToPdf_Gif.pdf"));
    std::cout << "ImageToPdf example finished." << std::endl << std::endl;
}*/