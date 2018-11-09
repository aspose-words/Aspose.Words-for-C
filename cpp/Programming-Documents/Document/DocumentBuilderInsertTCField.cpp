#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>

using namespace Aspose::Words;

void DocumentBuilderInsertTCField()
{
    std::cout << "DocumentBuilderInsertTCField example started." << std::endl;
    // ExStart:DocumentBuilderInsertTCField
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();

    // Create a document builder to insert content with.
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Insert a TC field at the current document builder position.
    builder->InsertField(u"TC \"Entry Text\" \\f t");

    System::String outputPath = dataDir + GetOutputFilePath(u"DocumentBuilderInsertTCField.doc");
    doc->Save(outputPath);
    // ExEnd:DocumentBuilderInsertTCField
    std::cout << "TC field inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "DocumentBuilderInsertTCField example finished." << std::endl << std::endl;
}