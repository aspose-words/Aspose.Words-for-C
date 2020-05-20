#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void LoadAndSaveToDisk()
{
    std::cout << "LoadAndSaveToDisk example started." << std::endl;
    // ExStart:LoadAndSave
    // ExStart:OpenDocument
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();
    // Load the document from the absolute path on disk.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.docx");
    // ExEnd:OpenDocument
    System::String outputPath = outputDataDir + u"LoadAndSaveToDisk.docx";
    // Save the document as DOC document.");
    doc->Save(outputPath);
    // ExEnd:LoadAndSave
    std::cout << "Existing document loaded and saved successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "LoadAndSaveToDisk example finished." << std::endl << std::endl;
}