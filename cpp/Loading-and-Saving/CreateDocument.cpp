#include "stdafx.h"
#include "examples.h"

#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void CreateDocument()
{
    std::cout << "CreateDocument example started." << std::endl;
    // ExStart:CreateDocument
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();

    // Initialize a Document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    
    // Use a document builder to add content to the document.
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"Hello World!");

    System::String outputPath = dataDir + GetOutputFilePath(u"CreateDocument.doc");
    // Save the document to disk.
    doc->Save(outputPath);
    // ExEnd:CreateDocument
    std::cout << "Document created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CreateDocument example finished." << std::endl << std::endl;
}