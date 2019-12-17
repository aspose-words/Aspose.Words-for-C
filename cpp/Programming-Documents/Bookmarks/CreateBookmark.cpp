#include "stdafx.h"
#include "examples.h"

#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Bookmarks/BookmarkStart.h>
#include <Model/Bookmarks/BookmarkEnd.h>
#include <Model/Saving/PdfSaveOptions.h>
#include <Model/Saving/OutlineOptions.h>
#include <Model/Saving/BookmarksOutlineLevelCollection.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

void CreateBookmark()
{
    std::cout << "CreateBookmark example started." << std::endl;
    // ExStart:CreateBookmark
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithBookmarks();

    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    builder->StartBookmark(u"My Bookmark");
    builder->Writeln(u"Text inside a bookmark.");

    builder->StartBookmark(u"Nested Bookmark");
    builder->Writeln(u"Text inside a NestedBookmark.");
    builder->EndBookmark(u"Nested Bookmark");

    builder->Writeln(u"Text after Nested Bookmark.");
    builder->EndBookmark(u"My Bookmark");

    System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
    options->get_OutlineOptions()->get_BookmarksOutlineLevels()->Add(u"My Bookmark", 1);
    options->get_OutlineOptions()->get_BookmarksOutlineLevels()->Add(u"Nested Bookmark", 2);

    System::String outputPath = outputDataDir + u"CreateBookmark.pdf";
    doc->Save(outputPath, options);
    // ExEnd:CreateBookmark
    std::cout << "Bookmark created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CreateBookmark example finished." << std::endl << std::endl;
}
