#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>

using namespace Aspose::Words;

void DocumentBuilderInsertHorizontalRule()
{
    std::cout << "DocumentBuilderInsertHorizontalRule example started." << std::endl;
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

    // ExStart:DocumentBuilderInsertHorizontalRule
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    builder->Writeln(u"Insert a horizontal rule shape into the document.");
    builder->InsertHorizontalRule();

    System::String outputPath = outputDataDir + u"DocumentBuilderInsertHorizontalRule.doc";
    doc->Save(outputPath);
    // ExEnd:DocumentBuilderInsertHorizontalRule
    std::cout << "Horizontal rule is inserted into document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "DocumentBuilderInsertHorizontalRule example finished." << std::endl << std::endl;
}