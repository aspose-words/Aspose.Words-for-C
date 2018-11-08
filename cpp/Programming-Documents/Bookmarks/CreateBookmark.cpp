#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Saving/DocSaveOptions.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Bookmarks/BookmarkStart.h>
#include <Model/Bookmarks/BookmarkEnd.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

void CreateBookmark()
{
    std::cout << "CreateBookmark example started." << std::endl;
    // ExStart:CreateBookmark
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithBookmarks();
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    builder->StartBookmark(u"My Bookmark");
    builder->Writeln(u"Text inside a bookmark.");
    
    builder->StartBookmark(u"Nested Bookmark");
    builder->Writeln(u"Text inside a NestedBookmark.");
    builder->EndBookmark(u"Nested Bookmark");
    
    builder->Writeln(u"Text after Nested Bookmark.");
    builder->EndBookmark(u"My Bookmark");
    
    System::SharedPtr<DocSaveOptions> options = System::MakeObject<DocSaveOptions>();
    System::String outputPath = dataDir + GetOutputFilePath(u"CreateBookmark.doc");
    doc->Save(outputPath, options);
    // ExEnd:CreateBookmark
    std::cout << "Bookmark created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CreateBookmark example finished." << std::endl << std::endl;
}
