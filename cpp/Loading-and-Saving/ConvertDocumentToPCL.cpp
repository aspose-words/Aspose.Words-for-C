#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/PclSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

void ConvertDocumentToPCL()
{
    std::cout << "ConvertDocumentToPCL example started." << std::endl;
    // ExStart:ConvertDocumentToPCL
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

    // Load the document from disk.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Test File (docx).docx");

    System::SharedPtr<PclSaveOptions> saveOptions = System::MakeObject<PclSaveOptions>();

    saveOptions->set_SaveFormat(SaveFormat::Pcl);
    saveOptions->set_RasterizeTransformedElements(false);

    // Export the document as an PCL file.
    doc->Save(outputDataDir + u"ConvertDocumentToPCL.pcl", saveOptions);
    // ExEnd:ConvertDocumentToPCL
    std::cout << "ConvertDocumentToPCL example finished." << std::endl << std::endl;
}