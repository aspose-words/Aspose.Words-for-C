#include "stdafx.h"
#include "examples.h"

#include <Model/Tables/Table.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/Cell.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Bookmarks/BookmarkStart.h>
#include <Model/Bookmarks/BookmarkEnd.h>


using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

void BookmarkTable()
{
    std::cout << "BookmarkTable example started." << std::endl;
    // ExStart:BookmarkTable
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithBookmarks();

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
    std::cout << "BookmarkTable example finished." << std::endl << std::endl;
}
