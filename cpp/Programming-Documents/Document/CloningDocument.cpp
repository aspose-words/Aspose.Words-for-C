#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void CloningDocument()
{
    std::cout << "CloningDocument example started." << std::endl;
    // ExStart:CloningDocument
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();

    // Load the document from disk.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");

    System::SharedPtr<Document> clone = doc->Clone();

    System::String outputPath = dataDir + GetOutputFilePath(u"CloningDocument.doc");

    // Save the document to disk.
    clone->Save(outputPath);
    // ExEnd:CloningDocument
    std::cout << "Document cloned successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CloningDocument example finished." << std::endl << std::endl;
}
