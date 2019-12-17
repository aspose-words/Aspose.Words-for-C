#include "stdafx.h"
#include "examples.h"

#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Document/BreakType.h>

using namespace Aspose::Words;

void DocumentBuilderInsertBreak()
{
    std::cout << "DocumentBuilderInsertBreak example started." << std::endl;
    // ExStart:DocumentBuilderInsertBreak
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    builder->Writeln(u"This is page 1.");
    builder->InsertBreak(BreakType::PageBreak);

    builder->Writeln(u"This is page 2.");
    builder->InsertBreak(BreakType::PageBreak);

    builder->Writeln(u"This is page 3.");
    System::String outputPath = outputDataDir + u"DocumentBuilderInsertBreak.doc";
    doc->Save(outputPath);
    // ExEnd:DocumentBuilderInsertBreak
    std::cout << "Page breaks inserted into a document using DocumentBuilder." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "DocumentBuilderInsertBreak example finished." << std::endl << std::endl;
}
