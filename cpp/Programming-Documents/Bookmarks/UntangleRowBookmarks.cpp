#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <Model/Text/Range.h>
#include <Model/Text/Paragraph.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/Cell.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/CompositeNode.h>
#include <Model/Document/Document.h>
#include <Model/Bookmarks/BookmarkStart.h>
#include <Model/Bookmarks/BookmarkEnd.h>
#include <Model/Bookmarks/BookmarkCollection.h>
#include <Model/Bookmarks/Bookmark.h>


using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace
{
    void Untangle(const System::SharedPtr<Document>& doc)
    {
        for(System::SharedPtr<Bookmark>  bookmark : System::IterateOver(doc->get_Range()->get_Bookmarks()))
        {
            // Get the parent row of both the bookmark and bookmark end node.
            System::SharedPtr<Row> row1 = System::DynamicCast<Row>(bookmark->get_BookmarkStart()->GetAncestor(NodeType::Row));
            System::SharedPtr<Row> row2 = System::DynamicCast<Row>(bookmark->get_BookmarkEnd()->GetAncestor(NodeType::Row));

            // If both rows are found okay and the bookmark start and end are contained
            // In adjacent rows, then just move the bookmark end node to the end
            // Of the last paragraph in the last cell of the top row.
            if ((row1 != nullptr) && (row2 != nullptr) && (row1->get_NextSibling() == row2))
            {
                row1->get_LastCell()->get_LastParagraph()->AppendChild(bookmark->get_BookmarkEnd());
            }
        }
    }

    void DeleteRowByBookmark(const System::SharedPtr<Document>& doc, const System::String& bookmarkName)
    {
        // Find the bookmark in the document. Exit if cannot find it.
        auto bookmark = doc->get_Range()->get_Bookmarks()->idx_get(bookmarkName);
        if (bookmark == nullptr)
        {
            return;
        }

        // Get the parent row of the bookmark. Exit if the bookmark is not in a row.
        auto row = System::DynamicCast<Row>(bookmark->get_BookmarkStart()->GetAncestor(NodeType::Row));
        if (row == nullptr)
        {
            return;
        }

        // Remove the row.
        row->Remove();
    }
}

void UntangleRowBookmarks()
{
    std::cout << "UntangleRowBookmarks example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithBookmarks();

    // Load a document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestDefect1352.doc");

    // This perform the custom task of putting the row bookmark ends into the same row with the bookmark starts.
    Untangle(doc);

    // Now we can easily delete rows by a bookmark without damaging any other row's bookmarks.
    DeleteRowByBookmark(doc, u"ROW2");

    // This is just to check that the other bookmark was not damaged.
    if (doc->get_Range()->get_Bookmarks()->idx_get(u"ROW1")->get_BookmarkEnd() == nullptr)
    {
        throw System::Exception(u"Wrong, the end of the bookmark was deleted.");
    }

    System::String outputPath = dataDir + GetOutputFilePath(u"UntangleRowBookmarks.doc");
    // Save the finished document.
    doc->Save(outputPath);

    std::cout << "Row bookmark untangled successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "UntangleRowBookmarks example finished." << std::endl << std::endl;
}
