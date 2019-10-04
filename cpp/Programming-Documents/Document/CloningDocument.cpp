#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void CloningDocument()
{
    std::cout << "CloningDocument example started." << std::endl;
    // ExStart:CloningDocument
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

    // Load the document from disk.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");

    System::SharedPtr<Document> clone = doc->Clone();

    System::String outputPath = outputDataDir + u"CloningDocument.doc";

    // Save the document to disk.
    clone->Save(outputPath);
    // ExEnd:CloningDocument
    std::cout << "Document cloned successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CloningDocument example finished." << std::endl << std::endl;
}
