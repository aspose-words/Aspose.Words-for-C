#include "../examples.h"

#include "system/string.h"
#include "Model/Saving/SaveOutputParameters.h"
#include "Model/Document/Document.h"

using namespace Aspose::Words;

void LoadAndSaveToDisk()
{
    // ExStart:LoadAndSave
    // ExStart:OpenDocument
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();
    // Load the document from the absolute path on disk.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");
    // ExEnd:OpenDocument
    dataDir = dataDir + GetOutputFilePath(u"LoadAndSaveToDisk.doc");
    // Save the document as DOC document.");
    doc->Save(dataDir);
    // ExEnd:LoadAndSave
    std::cout << "\nExisting document loaded and saved successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}