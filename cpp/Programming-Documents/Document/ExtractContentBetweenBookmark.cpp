#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/collections/list.h>
#include <system/collections/ilist.h>
#include <Model/Text/Range.h>
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/Node.h>
#include <Model/Document/Document.h>
#include <Model/Bookmarks/BookmarkStart.h>
#include <Model/Bookmarks/BookmarkEnd.h>
#include <Model/Bookmarks/BookmarkCollection.h>
#include <Model/Bookmarks/Bookmark.h>

#include "Common.h"

using namespace Aspose::Words;


void ExtractContentBetweenBookmark()
{
    // ExStart:ExtractContentBetweenBookmark
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    
    System::SharedPtr<Section> section = doc->get_Sections()->idx_get(0);
    section->get_PageSetup()->set_LeftMargin(70.85);
    
    // Retrieve the bookmark from the document.
    System::SharedPtr<Bookmark> bookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"Bookmark1");
    
    // We use the BookmarkStart and BookmarkEnd nodes as markers.
    System::SharedPtr<BookmarkStart> bookmarkStart = bookmark->get_BookmarkStart();
    System::SharedPtr<BookmarkEnd> bookmarkEnd = bookmark->get_BookmarkEnd();
    
    // Firstly extract the content between these nodes including the bookmark. 
    auto extractedNodesInclusive = ExtractContent(bookmarkStart, bookmarkEnd, true);
    System::SharedPtr<Document> dstDoc = GenerateDocument(doc, extractedNodesInclusive);
    dstDoc->Save(dataDir + GetOutputFilePath(u"ExtractContentBetweenBookmark.Inclusive.doc"));
    
    // Secondly extract the content between these nodes this time without including the bookmark.
    auto extractedNodesExclusive = ExtractContent(bookmarkStart, bookmarkEnd, false);
    dstDoc = GenerateDocument(doc, extractedNodesExclusive);
    dataDir += GetOutputFilePath(u"ExtractContentBetweenBookmark.Exclusive.doc");
    dstDoc->Save(dataDir);
    // ExEnd:ExtractContentBetweenBookmark
    std::cout << "\nExtracted content between bookmarks successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
