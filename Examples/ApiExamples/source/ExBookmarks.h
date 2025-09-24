#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/shared_ptr.h>
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
    void CreateUpdateAndPrintBookmarks();
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


