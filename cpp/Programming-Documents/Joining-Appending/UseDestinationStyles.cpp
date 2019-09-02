#include "stdafx.h"
#include "examples.h"

#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void UseDestinationStyles()
{
    std::cout << "UseDestinationStyles example started." << std::endl;
    // ExStart:UseDestinationStyles
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_JoiningAndAppending();
    System::String outputDataDir = GetOutputDataDir_JoiningAndAppending();

    // Load the documents to join.
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Source.doc");

    // Append the source document using the styles of the destination document.
    dstDoc->AppendDocument(srcDoc, ImportFormatMode::UseDestinationStyles);

    // Save the joined document to disk.
    System::String outputPath = outputDataDir + u"UseDestinationStyles.doc";
    dstDoc->Save(outputPath);
    // ExEnd:UseDestinationStyles
    std::cout << "Document appended successfully with use destination styles option." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "UseDestinationStyles example finished." << std::endl << std::endl;
}

