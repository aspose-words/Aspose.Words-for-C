#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Document/BreakType.h>

using namespace Aspose::Words;

void DocumentBuilderInsertBreak()
{
    // ExStart:DocumentBuilderInsertBreak
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    builder->Writeln(u"This is page 1.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    builder->Writeln(u"This is page 2.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    builder->Writeln(u"This is page 3.");
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderInsertBreak.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderInsertBreak
    std::cout << "\nPage breaks inserted into a document using DocumentBuilder.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
