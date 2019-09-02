#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include "Model/Importing/ImportFormatMode.h"

using namespace Aspose::Words;

void AppendDocuments()
{
    std::cout << "AppendDocuments example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_QuickStart();
    System::String outputDataDir = GetOutputDataDir_QuickStart();
    // Load the destination and source documents from disk.
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Source.doc");

    // Append the source document to the destination document while keeping the original formatting of the source document.
    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

    System::String outputPath = outputDataDir + u"AppendDocuments.doc";
    dstDoc->Save(outputPath);
    std::cout << "Document appended successfully." << std::endl << "File saved at" << outputPath.ToUtf8String() << '\n';
    std::cout << "AppendDocuments example finished." << std::endl;
}
