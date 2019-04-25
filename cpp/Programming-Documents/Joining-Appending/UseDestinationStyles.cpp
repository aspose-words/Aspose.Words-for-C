#include "stdafx.h"
#include "examples.h"

#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void UseDestinationStyles()
{
    std::cout << "UseDestinationStyles example started." << std::endl;
    // ExStart:UseDestinationStyles
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();

    // Load the documents to join.
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");

    // Append the source document using the styles of the destination document.
    dstDoc->AppendDocument(srcDoc, ImportFormatMode::UseDestinationStyles);

    // Save the joined document to disk.
    System::String outputPath = dataDir + GetOutputFilePath(u"UseDestinationStyles.doc");
    dstDoc->Save(outputPath);
    // ExEnd:UseDestinationStyles
    std::cout << "Document appended successfully with use destination styles option." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "UseDestinationStyles example finished." << std::endl << std::endl;
}

