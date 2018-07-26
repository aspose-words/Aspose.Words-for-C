#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
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
    // ExStart:BookmarkTable
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithBookmarks();
    
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
    
    dataDir = dataDir + GetOutputFilePath(u"BookmarkTable.doc");
    doc->Save(dataDir);
    // ExEnd:BookmarkTable
    std::cout << "\nTable bookmarked successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
