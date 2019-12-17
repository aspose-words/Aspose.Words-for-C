#include "stdafx.h"
#include "examples.h"

#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void KeepSourceFormatting()
{
    std::cout << "KeepSourceFormatting example started." << std::endl;
    // ExStart:KeepSourceFormatting
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_JoiningAndAppending();
    System::String outputDataDir = GetOutputDataDir_JoiningAndAppending();
    // Load the documents to join.
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Source.doc");
    
    // Keep the formatting from the source document when appending it to the destination document.
    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);
    
    // Save the joined document to disk.
    System::String outputPath = outputDataDir + u"KeepSourceFormatting.doc";
    dstDoc->Save(outputPath);
    // ExEnd:KeepSourceFormatting
    std::cout << "Document appended successfully with keep source formatting option." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "KeepSourceFormatting example finished." << std::endl << std::endl;
}