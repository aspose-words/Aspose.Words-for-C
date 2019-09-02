#include "stdafx.h"
#include "examples.h"

#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Bookmarks/BookmarkStart.h>
#include <Model/Bookmarks/BookmarkEnd.h>

using namespace Aspose::Words;

void DocumentBuilderInsertBookmark()
{
    std::cout << "DocumentBuilderInsertBookmark example started." << std::endl;
    // ExStart:DocumentBuilderInsertBookmark
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    builder->StartBookmark(u"FineBookmark");
    builder->Writeln(u"This is just a fine bookmark.");
    builder->EndBookmark(u"FineBookmark");

    System::String outputPath = outputDataDir + u"DocumentBuilderInsertBookmark.doc";
    doc->Save(outputPath);
    // ExEnd:DocumentBuilderInsertBookmark
    std::cout << "Bookmark using DocumentBuilder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "DocumentBuilderInsertBookmark example finished." << std::endl << std::endl;
}
