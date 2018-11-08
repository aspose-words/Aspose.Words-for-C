#include "stdafx.h"
#include "examples.h"

#include "system/string.h"
#include "Model/Saving/SaveOutputParameters.h"
#include "Model/Document/Document.h"

using namespace Aspose::Words;

void LoadAndSaveToDisk()
{
    std::cout << "LoadAndSaveToDisk example started." << std::endl;
    // ExStart:LoadAndSave
    // ExStart:OpenDocument
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();
    // Load the document from the absolute path on disk.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");
    // ExEnd:OpenDocument
    System::String outputPath = dataDir + GetOutputFilePath(u"LoadAndSaveToDisk.doc");
    // Save the document as DOC document.");
    doc->Save(outputPath);
    // ExEnd:LoadAndSave
    std::cout << "Existing document loaded and saved successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "LoadAndSaveToDisk example finished." << std::endl << std::endl;
}