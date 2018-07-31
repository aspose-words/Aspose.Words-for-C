#include "stdafx.h"
#include "examples.h"

#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void CreateDocument()
{
    // ExStart:CreateDocument
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();
    
    // Initialize a Document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    
    // Use a document builder to add content to the document.
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"Hello World!");
    
    dataDir = dataDir + GetOutputFilePath(u"CreateDocument.doc");
    // Save the document to disk.
    doc->Save(dataDir);
    
    // ExEnd:CreateDocument
    
    std::cout << "\nDocument created successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}