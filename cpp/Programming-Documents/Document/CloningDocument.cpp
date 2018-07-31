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
    // ExStart:CloningDocument
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    
    // Load the document from disk.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    
    System::SharedPtr<Document> clone = doc->Clone();
    
    dataDir = dataDir + GetOutputFilePath(u"CloningDocument.doc");
    
    // Save the document to disk.
    clone->Save(dataDir);
    // ExEnd:CloningDocument
    std::cout << "\nDocument cloned successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
