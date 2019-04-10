#include "stdafx.h"
#include "examples.h"

#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void SimpleAppendDocument()
{
    std::cout << "SimpleAppendDocument example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");

    // Append the source document to the destination document using no extra options.
    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

    System::String outputPath = dataDir + GetOutputFilePath(u"SimpleAppendDocument.doc");
    dstDoc->Save(outputPath);
    std::cout << "Simple document append." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "SimpleAppendDocument example finished." << std::endl << std::endl;
}
