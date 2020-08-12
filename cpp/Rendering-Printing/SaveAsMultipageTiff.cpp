#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/TiffCompression.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

void SaveAsMultipageTiff()
{
    std::cout << "SaveAsMultipageTiff example started." << std::endl;
    // ExStart:SaveAsMultipageTiff
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    // Open the document.
    auto doc = System::MakeObject<Document>(inputDataDir + u"TestFile Multipage TIFF.doc");

    // ExStart:SaveAsTIFF
    // Save the document as multipage TIFF.
    doc->Save(outputDataDir + u"SaveAsMultipageTiff.WithoutOptions.tiff");
    // ExEnd:SaveAsTIFF
    // ExStart:SaveAsTIFFUsingImageSaveOptions
    // Create an ImageSaveOptions object to pass to the Save method
    System::SharedPtr<ImageSaveOptions> options = System::MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
    options->set_PageIndex(0);
    options->set_PageCount(2);
    options->set_TiffCompression(TiffCompression::Ccitt4);
    options->set_Resolution(160.0f);
    System::String outputPath = outputDataDir + u"SaveAsMultipageTiff.WithOptions.tiff";
    doc->Save(outputPath, options);
    // ExEnd:SaveAsTIFFUsingImageSaveOptions
    // ExEnd:SaveAsMultipageTiff
    std::cout << "Document saved as multi-page TIFF successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "SaveAsMultipageTiff example finished." << std::endl << std::endl;
}