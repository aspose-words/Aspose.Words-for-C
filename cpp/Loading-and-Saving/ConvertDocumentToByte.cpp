﻿#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void ConvertDocumentToByte()
{
    std::cout << "ConvertDocumentToByte example started." << std::endl;
    // ExStart:ConvertDocumentToByte
    // The path to the documents directory.
    System::String dataDir = GetInputDataDir_LoadingAndSaving();

    // Load the document from disk.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Test File (doc).doc");

    // Create a new memory stream.
    System::SharedPtr<System::IO::MemoryStream> outStream = System::MakeObject<System::IO::MemoryStream>();
    // Save the document to stream.
    doc->Save(outStream, SaveFormat::Docx);

    // Convert the document to byte form.
    System::ArrayPtr<uint8_t> docBytes = outStream->ToArray();

    // The bytes are now ready to be stored/transmitted.
    // Now reverse the steps to load the bytes back into a document object.
    System::SharedPtr<System::IO::MemoryStream> inStream = System::MakeObject<System::IO::MemoryStream>(docBytes);

    // Load the stream into a new document object.
    System::SharedPtr<Document> loadDoc = System::MakeObject<Document>(inStream);
    // ExEnd:ConvertDocumentToByte
    std::cout << "Document converted to byte array successfully." << std::endl;
    std::cout << "ConvertDocumentToByte example finished." << std::endl << std::endl;
}
