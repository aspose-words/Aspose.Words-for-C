#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/console.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void KeepSourceFormatting()
{
    std::cout << "KeepSourceFormatting example started." << std::endl;
    // ExStart:KeepSourceFormatting
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();
    // Load the documents to join.
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");
    
    // Keep the formatting from the source document when appending it to the destination document.
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting);
    
    // Save the joined document to disk.
    System::String outputPath = dataDir + GetOutputFilePath(u"KeepSourceFormatting.doc");
    dstDoc->Save(outputPath);
    // ExEnd:KeepSourceFormatting
    std::cout << "Document appended successfully with keep source formatting option." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "KeepSourceFormatting example finished." << std::endl << std::endl;
}