#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>


using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace
{
    void InsertBookmarkTable(System::String const &outputDataDir)
    {
        // ExStart:BookmarkTable
        // Create empty document
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::SharedPtr<Table> table = builder->StartTable();

        // Insert a cell
        builder->InsertCell();

        // Start bookmark here after calling InsertCell
        builder->StartBookmark(u"MyBookmark");

        builder->Write(u"This is row 1 cell 1");

        // Insert a cell
        builder->InsertCell();
        builder->Write(u"This is row 1 cell 2");

        builder->EndRow();

        // Insert a cell
        builder->InsertCell();
        builder->Writeln(u"This is row 2 cell 1");

        // Insert a cell
        builder->InsertCell();
        builder->Writeln(u"This is row 2 cell 2");

        builder->EndRow();

        builder->EndTable();
        // End of bookmark
        builder->EndBookmark(u"MyBookmark");

        System::String outputPath = outputDataDir + u"BookmarkTable.doc";
        doc->Save(outputPath);
        // ExEnd:BookmarkTable
        std::cout << "Table bookmarked successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void BookmarkTableColumns(System::String const &inputDataDir)
    {
        // ExStart:BookmarkTableColumns
        // Create empty document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"BookmarkTable.doc");
        for (System::SharedPtr<Bookmark> bookmark : System::IterateOver<Bookmark>(doc->get_Range()->get_Bookmarks()))
        {
            std::cout << "Bookmark: " << bookmark->get_Name().ToUtf8String() << (bookmark->get_IsColumn() ? " (Column)" : "") << std::endl;
            if (bookmark->get_IsColumn())
            {
                System::SharedPtr<Row> row = System::StaticCast<Row>(bookmark->get_BookmarkStart()->GetAncestor(NodeType::Row));
                if (row != nullptr && bookmark->get_FirstColumn() < row->get_Cells()->get_Count())
                {
                    std::cout << row->get_Cells()->idx_get(bookmark->get_FirstColumn())->GetText().TrimEnd(ControlChar::CellChar).ToUtf8String() << std::endl;
                }
            }
        }
        // ExEnd:BookmarkTableColumns
    }
}

void BookmarkTable()
{
    std::cout << "BookmarkTable example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithBookmarks();
    System::String outputDataDir = GetOutputDataDir_WorkingWithBookmarks();
    InsertBookmarkTable(outputDataDir);
    BookmarkTableColumns(inputDataDir);
    std::cout << "BookmarkTable example finished." << std::endl << std::endl;
}
