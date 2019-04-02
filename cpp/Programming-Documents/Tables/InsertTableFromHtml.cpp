#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>

using namespace Aspose::Words;

void InsertTableFromHtml()
{
    std::cout << "InsertTableFromHtml example started." << std::endl;
    // ExStart:InsertTableFromHtml
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithTables();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Insert the table from HTML. Note that AutoFitSettings does not apply to tables
    // Inserted from HTML.
    builder->InsertHtml(u"<table><tr><td>Row 1, Cell 1</td><td>Row 1, Cell 2</td></tr><tr><td>Row 2, Cell 2</td><td>Row 2, Cell 2</td></tr></table>");

    System::String outputPath = dataDir + GetOutputFilePath(u"InsertTableFromHtml.doc");
    // Save the document to disk.
    doc->Save(outputPath);
    // ExEnd:InsertTableFromHtml
    std::cout << "Table inserted successfully from html." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertTableFromHtml example finished." << std::endl << std::endl;
}