#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <Model/Text/Range.h>
#include <Model/Document/Document.h>
#include <Model/Bookmarks/BookmarkCollection.h>
#include <Model/Bookmarks/Bookmark.h>

using namespace Aspose::Words;

void BookmarkNameAndText()
{
    // ExStart:BookmarkNameAndText
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithBookmarks();
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Bookmark.doc");
    
    // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
    System::SharedPtr<Bookmark> bookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark");
    
    // Get the name and text of the bookmark.
    System::String name = bookmark->get_Name();
    System::String text = bookmark->get_Text();
    
    // Set the name and text of the bookmark.
    bookmark->set_Name(u"RenamedBookmark");
    bookmark->set_Text(u"This is a new bookmarked text.");
    // ExEnd:BookmarkNameAndText
    std::cout << "\nBookmark text and name get and set successfully. " << dataDir.ToUtf8String() << '\n';
}
