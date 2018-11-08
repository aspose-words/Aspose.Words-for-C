#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>

using namespace Aspose::Words;

void DocumentBuilderInsertTOC()
{
    std::cout << "DocumentBuilderInsertTOC example started." << std::endl;
    // ExStart:DocumentBuilderInsertTOC
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Insert a table of contents at the beginning of the document.
    builder->InsertTableOfContents(u"\\o \"1-3\" \\h \\z \\u");

    // The newly inserted table of contents will be initially empty.
    // It needs to be populated by updating the fields in the document.
    // ExStart:UpdateFields
    doc->UpdateFields();
    // ExEnd:UpdateFields
    System::String outputPath = dataDir + GetOutputFilePath(u"DocumentBuilderInsertTOC.doc");
    doc->Save(outputPath);
    // ExEnd:DocumentBuilderInsertTOC
    std::cout << "Table of contents field inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "DocumentBuilderInsertTOC example finished." << std::endl << std::endl;
}