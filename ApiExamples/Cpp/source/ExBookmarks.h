#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <system/collections/ienumerator.h>
#include <system/details/dispose_guard.h>
#include <system/enumerator_adapter.h>
#include <system/func.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace ApiExamples {

class ExBookmarks : public ApiExampleBase
{
public:
    void Insert()
    {
        //ExStart
        //ExFor:Bookmark.Name
        //ExSummary:Shows how to insert a bookmark.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // A valid bookmark has a name, a BookmarkStart, and a BookmarkEnd node.
        // Any whitespace in the names of bookmarks will be converted to underscores if we open the saved document with Microsoft Word.
        // If we highlight the bookmark's name in Microsoft Word via Insert -> Links -> Bookmark, and press "Go To",
        // the cursor will jump to the text enclosed between the BookmarkStart and BookmarkEnd nodes.
        builder->StartBookmark(u"My Bookmark");
        builder->Write(u"Contents of MyBookmark.");
        builder->EndBookmark(u"My Bookmark");

        // Bookmarks are stored in this collection.
        ASSERT_EQ(u"My Bookmark", doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Name());

        doc->Save(ArtifactsDir + u"Bookmarks.Insert.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Bookmarks.Insert.docx");

        ASSERT_EQ(u"My Bookmark", doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Name());
    }

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
    void CreateUpdateAndPrintBookmarks()
    {
        // Create a document with three bookmarks, then use a custom document visitor implementation to print their contents.
        SharedPtr<Document> doc = CreateDocumentWithBookmarks(3);
        SharedPtr<BookmarkCollection> bookmarks = doc->get_Range()->get_Bookmarks();
        ASSERT_EQ(3, bookmarks->get_Count());
        //ExSkip

        PrintAllBookmarkInfo(bookmarks);

        // Bookmarks can be accessed in the bookmark collection by index or name, and their names can be updated.
        bookmarks->idx_get(0)->set_Name(String::Format(u"{0}_NewName", bookmarks->idx_get(0)->get_Name()));
        bookmarks->idx_get(u"MyBookmark_2")->set_Text(String::Format(u"Updated text contents of {0}", bookmarks->idx_get(1)->get_Name()));

        // Print all bookmarks again to see updated values.
        PrintAllBookmarkInfo(bookmarks);
    }

protected:
    /// <summary>
    /// Create a document with a given number of bookmarks.
    /// </summary>
    static SharedPtr<Document> CreateDocumentWithBookmarks(int numberOfBookmarks)
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        for (int i = 1; i <= numberOfBookmarks; i++)
        {
            String bookmarkName = String(u"MyBookmark_") + i;

            builder->Write(u"Text before bookmark.");
            builder->StartBookmark(bookmarkName);
            builder->Write(String::Format(u"Text inside {0}.", bookmarkName));
            builder->EndBookmark(bookmarkName);
            builder->Writeln(u"Text after bookmark.");
        }

        return doc;
    }

    /// <summary>
    /// Use an iterator and a visitor to print info of every bookmark in the collection.
    /// </summary>
    static void PrintAllBookmarkInfo(SharedPtr<BookmarkCollection> bookmarks)
    {
        auto bookmarkVisitor = MakeObject<ExBookmarks::BookmarkInfoPrinter>();

        // Get each bookmark in the collection to accept a visitor that will print its contents.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Bookmark>>> enumerator = bookmarks->GetEnumerator();
            while (enumerator->MoveNext())
            {
                SharedPtr<Bookmark> currentBookmark = enumerator->get_Current();

                if (currentBookmark != nullptr)
                {
                    currentBookmark->get_BookmarkStart()->Accept(bookmarkVisitor);
                    currentBookmark->get_BookmarkEnd()->Accept(bookmarkVisitor);

                    std::cout << currentBookmark->get_BookmarkStart()->GetText() << std::endl;
                }
            }
        }
    }

public:
    /// <summary>
    /// Prints contents of every visited bookmark to the console.
    /// </summary>
    class BookmarkInfoPrinter : public DocumentVisitor
    {
    public:
        VisitorAction VisitBookmarkStart(SharedPtr<BookmarkStart> bookmarkStart) override
        {
            std::cout << "BookmarkStart name: \"" << bookmarkStart->get_Name() << "\", Contents: \"" << bookmarkStart->get_Bookmark()->get_Text() << "\""
                      << std::endl;
            return VisitorAction::Continue;
        }

        VisitorAction VisitBookmarkEnd(SharedPtr<BookmarkEnd> bookmarkEnd) override
        {
            std::cout << "BookmarkEnd name: \"" << bookmarkEnd->get_Name() << "\"" << std::endl;
            return VisitorAction::Continue;
        }
    };
    //ExEnd

    void TableColumnBookmarks()
    {
        //ExStart
        //ExFor:Bookmark.IsColumn
        //ExFor:Bookmark.FirstColumn
        //ExFor:Bookmark.LastColumn
        //ExSummary:Shows how to get information about table column bookmarks.
        auto doc = MakeObject<Document>(MyDir + u"Table column bookmarks.doc");

        for (auto bookmark : System::IterateOver(doc->get_Range()->get_Bookmarks()))
        {
            // If a bookmark encloses columns of a table, it is a table column bookmark, and its IsColumn flag set to true.
            std::cout << "Bookmark: " << bookmark->get_Name() << (bookmark->get_IsColumn() ? String(u" (Column)") : String(u"")) << std::endl;
            if (bookmark->get_IsColumn())
            {
                auto row = System::DynamicCast_noexcept<Row>(bookmark->get_BookmarkStart()->GetAncestor(NodeType::Row));
                if (row != nullptr && bookmark->get_FirstColumn() < row->get_Cells()->get_Count())
                {
                    // Print the contents of the first and last columns enclosed by the bookmark.
                    std::cout << row->get_Cells()->idx_get(bookmark->get_FirstColumn())->GetText().TrimEnd(MakeArray<char16_t>({ControlChar::CellChar}))
                              << std::endl;
                    std::cout << row->get_Cells()->idx_get(bookmark->get_LastColumn())->GetText().TrimEnd(MakeArray<char16_t>({ControlChar::CellChar}))
                              << std::endl;
                }
            }
        }
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);

        SharedPtr<Bookmark> firstTableColumnBookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"FirstTableColumnBookmark");
        SharedPtr<Bookmark> secondTableColumnBookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"SecondTableColumnBookmark");

        ASSERT_TRUE(firstTableColumnBookmark->get_IsColumn());
        ASSERT_EQ(1, firstTableColumnBookmark->get_FirstColumn());
        ASSERT_EQ(3, firstTableColumnBookmark->get_LastColumn());

        ASSERT_TRUE(secondTableColumnBookmark->get_IsColumn());
        ASSERT_EQ(0, secondTableColumnBookmark->get_FirstColumn());
        ASSERT_EQ(3, secondTableColumnBookmark->get_LastColumn());
    }

    void Remove()
    {
        //ExStart
        //ExFor:BookmarkCollection.Clear
        //ExFor:BookmarkCollection.Count
        //ExFor:BookmarkCollection.Remove(Bookmark)
        //ExFor:BookmarkCollection.Remove(String)
        //ExFor:BookmarkCollection.RemoveAt
        //ExFor:Bookmark.Remove
        //ExSummary:Shows how to remove bookmarks from a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert five bookmarks with text inside their boundaries.
        for (int i = 1; i <= 5; i++)
        {
            String bookmarkName = String(u"MyBookmark_") + i;

            builder->StartBookmark(bookmarkName);
            builder->Write(String::Format(u"Text inside {0}.", bookmarkName));
            builder->EndBookmark(bookmarkName);
            builder->InsertBreak(BreakType::ParagraphBreak);
        }

        // This collection stores bookmarks.
        SharedPtr<BookmarkCollection> bookmarks = doc->get_Range()->get_Bookmarks();

        ASSERT_EQ(5, bookmarks->get_Count());

        // There are several ways of removing bookmarks.
        // 1 -  Calling the bookmark's Remove method:
        bookmarks->idx_get(u"MyBookmark_1")->Remove();

        ASSERT_FALSE(bookmarks->LINQ_Any([](SharedPtr<Bookmark> b) { return b->get_Name() == u"MyBookmark_1"; }));

        // 2 -  Passing the bookmark to the collection's Remove method:
        SharedPtr<Bookmark> bookmark = doc->get_Range()->get_Bookmarks()->idx_get(0);
        doc->get_Range()->get_Bookmarks()->Remove(bookmark);

        ASSERT_FALSE(bookmarks->LINQ_Any([](SharedPtr<Bookmark> b) { return b->get_Name() == u"MyBookmark_2"; }));

        // 3 -  Removing a bookmark from the collection by name:
        doc->get_Range()->get_Bookmarks()->Remove(u"MyBookmark_3");

        ASSERT_FALSE(bookmarks->LINQ_Any([](SharedPtr<Bookmark> b) { return b->get_Name() == u"MyBookmark_3"; }));

        // 4 -  Removing a bookmark at an index in the bookmark collection:
        doc->get_Range()->get_Bookmarks()->RemoveAt(0);

        ASSERT_FALSE(bookmarks->LINQ_Any([](SharedPtr<Bookmark> b) { return b->get_Name() == u"MyBookmark_4"; }));

        // We can clear the entire bookmark collection.
        bookmarks->Clear();

        // The text that was inside the bookmarks is still present in the document.
        ASSERT_EQ(0, bookmarks->get_Count());
        ASSERT_EQ(String(u"Text inside MyBookmark_1.\r") + u"Text inside MyBookmark_2.\r" + u"Text inside MyBookmark_3.\r" +
                      u"Text inside MyBookmark_4.\r" + u"Text inside MyBookmark_5.",
                  doc->GetText().Trim());
        //ExEnd
    }
};

} // namespace ApiExamples
