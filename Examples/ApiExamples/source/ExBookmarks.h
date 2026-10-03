#pragma once

#include <system/shared_ptr.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExBookmarks : public ApiExampleBase
{
    typedef ExBookmarks ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Prints contents of every visited bookmark to the console.
    /// </summary>
    class BookmarkInfoPrinter : public DocumentVisitor
    {
        typedef BookmarkInfoPrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        Aspose::Words::VisitorAction VisitBookmarkStart(System::SharedPtr<Aspose::Words::BookmarkStart> bookmarkStart) override;
        Aspose::Words::VisitorAction VisitBookmarkEnd(System::SharedPtr<Aspose::Words::BookmarkEnd> bookmarkEnd) override;
        
    };
    
    
public:

    void Insert();
    //ExStart
    //ExFor:Bookmark
    //ExFor:Bookmark.Name
    //ExFor:Bookmark.Text
    //ExFor:Bookmark.BookmarkStart
    //ExFor:Bookmark.BookmarkEnd
    //ExFor:BookmarkStart
    //ExFor:BookmarkStart.#ctor
    //ExFor:BookmarkEnd
    //ExFor:BookmarkEnd.#ctor
    //ExFor:BookmarkStart.Accept(DocumentVisitor)
    //ExFor:BookmarkEnd.Accept(DocumentVisitor)
    //ExFor:BookmarkStart.Bookmark
    //ExFor:BookmarkStart.GetText
    //ExFor:BookmarkStart.Name
    //ExFor:BookmarkEnd.Name
    //ExFor:BookmarkCollection
    //ExFor:BookmarkCollection.Item(Int32)
    //ExFor:BookmarkCollection.Item(String)
    //ExFor:BookmarkCollection.GetEnumerator
    //ExFor:Range.Bookmarks
    //ExFor:DocumentVisitor.VisitBookmarkStart
    //ExFor:DocumentVisitor.VisitBookmarkEnd
    //ExSummary:Shows how to add bookmarks and update their contents.
    void CreateUpdateAndPrintBookmarks();
    //ExEnd
    void TableColumnBookmarks();
    void Remove();
    
protected:

    /// <summary>
    /// Create a document with a given number of bookmarks.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Document> CreateDocumentWithBookmarks(int32_t numberOfBookmarks);
    /// <summary>
    /// Use an iterator and a visitor to print info of every bookmark in the collection.
    /// </summary>
    static void PrintAllBookmarkInfo(System::SharedPtr<Aspose::Words::BookmarkCollection> bookmarks);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


