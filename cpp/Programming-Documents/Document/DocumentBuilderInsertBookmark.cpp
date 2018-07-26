#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Bookmarks/BookmarkStart.h>
#include <Model/Bookmarks/BookmarkEnd.h>

using namespace Aspose::Words;


void DocumentBuilderInsertBookmark()
{
    // ExStart:DocumentBuilderInsertBookmark
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    builder->StartBookmark(u"FineBookmark");
    builder->Writeln(u"This is just a fine bookmark.");
    builder->EndBookmark(u"FineBookmark");
    
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderInsertBookmark.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderInsertBookmark
    std::cout << "\nBookmark using DocumentBuilder inserted successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
