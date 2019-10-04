#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>

using namespace Aspose::Words;

void BookmarkNameAndText()
{
    std::cout << "BookmarkNameAndText example started." << std::endl;
    // ExStart:BookmarkNameAndText
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithBookmarks();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Bookmark.doc");

    // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
    System::SharedPtr<Bookmark> bookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark");

    // Get the name and text of the bookmark.
    System::String name = bookmark->get_Name();
    System::String text = bookmark->get_Text();

    // Set the name and text of the bookmark.
    bookmark->set_Name(u"RenamedBookmark");
    bookmark->set_Text(u"This is a new bookmarked text.");
    // ExEnd:BookmarkNameAndText
    std::cout << "Bookmark text and name get and set successfully. " << std::endl;
    std::cout << "BookmarkNameAndText example finished." << std::endl << std::endl;
}
