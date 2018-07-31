#include "stdafx.h"
#include "examples.h"

#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/SaveFormat.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void ConvertDocumentToByte()
{
    // ExStart:ConvertDocumentToByte
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();
    
    // Load the document from disk.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Test File (doc).doc");
    
    // Create a new memory stream.
    System::SharedPtr<System::IO::MemoryStream> outStream = System::MakeObject<System::IO::MemoryStream>();
    // Save the document to stream.
    doc->Save(outStream, Aspose::Words::SaveFormat::Doc);
    
    // Convert the document to byte form.
    System::ArrayPtr<uint8_t> docBytes = outStream->ToArray();
    
    // The bytes are now ready to be stored/transmitted.
    
    // Now reverse the steps to load the bytes back into a document object.
    System::SharedPtr<System::IO::MemoryStream> inStream = System::MakeObject<System::IO::MemoryStream>(docBytes);
    
    // Load the stream into a new document object.
    System::SharedPtr<Document> loadDoc = System::MakeObject<Document>(inStream);
    // ExEnd:ConvertDocumentToByte
    
    std::cout << "\nDocument converted to byte array successfully.\n";
}
