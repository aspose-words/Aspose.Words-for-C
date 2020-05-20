#include "stdafx.h"
#include "examples.h"

#include <system/io/memory_stream.h>
#include <system/io/seekorigin.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/TiffCompression.h>

using namespace System::IO;
using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

void SaveDocumentToJPEG()
{
    std::cout << "SaveDocumentToJPEG example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();
    // ExStart:SaveDocumentToJPEG
    // Open the document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");
    
    // Save as a JPEG image file with default options
    System::String outputPathDefault = outputDataDir + u"Rendering.JpegDefaultOptions.jpg";
    doc->Save(outputPathDefault);

    // Save document to stream as a JPEG with default options
    System::SharedPtr<MemoryStream> docStream = new MemoryStream();
    doc->Save(docStream, SaveFormat::Jpeg);

    // Rewind the stream position back to the beginning, ready for use
    docStream->Seek(0, SeekOrigin::Begin);

    // Save document to a JPEG image with specified options.
    // Render the third page only and set the JPEG quality to 80%
    // In this case we need to pass the desired SaveFormat to the ImageSaveOptions constructor 
    // to signal what type of image to save as.
    System::SharedPtr<ImageSaveOptions> options = System::MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
    options->set_PageIndex(0);
    options->set_PageCount(1);
    options->set_JpegQuality(80);

    System::String outputPath = outputDataDir + u"Rendering.JpegCustomOptions.jpg";
    doc->Save(outputPath, options);
    // ExEnd:SaveDocumentToJPEG
    std::cout << "Document saved as JPEG successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "SaveDocumentToJPEG example finished." << std::endl << std::endl;
}