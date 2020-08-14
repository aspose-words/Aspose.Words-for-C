#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class Document; } }
namespace Aspose { namespace Words { class BookmarkCollection; } }
namespace Aspose { namespace Words { enum class VisitorAction; } }
namespace Aspose { namespace Words { class BookmarkStart; } }
namespace Aspose { namespace Words { class BookmarkEnd; } }

namespace ApiExamples {

class ExBookmarks : public ApiExampleBase
{
public:

    /// <summary>
    /// Visitor that prints bookmark information to the console.
    /// </summary>
    class BookmarkInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef BookmarkInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        Aspose::Words::VisitorAction VisitBookmarkStart(System::SharedPtr<Aspose::Words::BookmarkStart> bookmarkStart) override;
        Aspose::Words::VisitorAction VisitBookmarkEnd(System::SharedPtr<Aspose::Words::BookmarkEnd> bookmarkEnd) override;
        
    };
    
    
public:

    void CreateUpdateAndPrintBookmarks();
    void TableColumnBookmarks();
    void ClearBookmarks();
    void RemoveBookmarkFromBookmarkCollection();
    void ReplaceBookmarkUnderscoresWithWhitespaces();
    
protected:

    /// <summary>
    /// Create a document with bookmarks using the start and end nodes.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Document> CreateDocumentWithBookmarks();
    /// <summary>
    /// Use an iterator and a visitor to print info of every bookmark from within a document.
    /// </summary>
    static void PrintAllBookmarkInfo(System::SharedPtr<Aspose::Words::BookmarkCollection> bookmarks);
    
};

} // namespace ApiExamples


