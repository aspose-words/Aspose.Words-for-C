#include "stdafx.h"
#include "examples.h"

#include <Model/Document/CleanupOptions.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void CleansUnusedStylesandLists()
{
    std::cout << "CleansUnusedStylesandLists example started." << std::endl;
    // ExStart:CleansUnusedStylesandLists
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

    System::SharedPtr<CleanupOptions> cleanupoptions = System::MakeObject<CleanupOptions>();
    cleanupoptions->set_UnusedLists(false);
    cleanupoptions->set_UnusedStyles(true);

    // Cleans unused styles and lists from the document depending on given CleanupOptions. 
    doc->Cleanup(cleanupoptions);

    System::String outputPath = outputDataDir + u"CleansUnusedStylesandLists.docx";
    doc->Save(outputPath);
    // ExEnd:CleansUnusedStylesandLists
    std::cout << "All revisions accepted." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CleansUnusedStylesandLists example finished." << std::endl << std::endl;
}