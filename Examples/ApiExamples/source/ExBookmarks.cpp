// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBookmarks.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/func.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/collections/ienumerator.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>

#include "DocumentHelper.h"


using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(701186921u, ::Aspose::Words::ApiExamples::ExBookmarks::BookmarkInfoPrinter, ThisTypeBaseTypesInfo);

Aspose::Words::VisitorAction ExBookmarks::BookmarkInfoPrinter::VisitBookmarkStart(System::SharedPtr<Aspose::Words::BookmarkStart> bookmarkStart)
{
    std::cout << System::String::Format(u"BookmarkStart name: \"{0}\", Contents: \"{1}\"", bookmarkStart->get_Name(), bookmarkStart->get_Bookmark()->get_Text()) << std::endl;
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExBookmarks::BookmarkInfoPrinter::VisitBookmarkEnd(System::SharedPtr<Aspose::Words::BookmarkEnd> bookmarkEnd)
{
    std::cout << System::String::Format(u"BookmarkEnd name: \"{0}\"", bookmarkEnd->get_Name()) << std::endl;
    return Aspose::Words::VisitorAction::Continue;
}


RTTI_INFO_IMPL_HASH(1591165677u, ::Aspose::Words::ApiExamples::ExBookmarks, ThisTypeBaseTypesInfo);

System::SharedPtr<Aspose::Words::Document> ExBookmarks::CreateDocumentWithBookmarks(int32_t numberOfBookmarks)
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    for (int32_t i = 1; i <= numberOfBookmarks; i++)
    {
        System::String bookmarkName = System::String(u"MyBookmark_") + i;
        
        builder->Write(u"Text before bookmark.");
        builder->StartBookmark(bookmarkName);
        builder->Write(System::String::Format(u"Text inside {0}.", bookmarkName));
        builder->EndBookmark(bookmarkName);
        builder->Writeln(u"Text after bookmark.");
    }
    
    return doc;
}

void ExBookmarks::PrintAllBookmarkInfo(System::SharedPtr<Aspose::Words::BookmarkCollection> bookmarks)
{
    auto bookmarkVisitor = System::MakeObject<Aspose::Words::ApiExamples::ExBookmarks::BookmarkInfoPrinter>();
    
    // Get each bookmark in the collection to accept a visitor that will print its contents.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Bookmark>>> enumerator = bookmarks->GetEnumerator();
        while (enumerator->MoveNext())
        {
            System::SharedPtr<Aspose::Words::Bookmark> currentBookmark = enumerator->get_Current();
            
            if (currentBookmark != nullptr)
            {
                currentBookmark->get_BookmarkStart()->Accept(bookmarkVisitor);
                currentBookmark->get_BookmarkEnd()->Accept(bookmarkVisitor);
                
                std::cout << currentBookmark->get_BookmarkStart()->GetText() << std::endl;
            }
        }
    }
}


namespace gtest_test
{

class ExBookmarks : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExBookmarks> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExBookmarks>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExBookmarks> ExBookmarks::s_instance;

} // namespace gtest_test

void ExBookmarks::Insert()
{
    //ExStart
    //ExFor:Bookmark.Name
    //ExSummary:Shows how to insert a bookmark.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // A valid bookmark has a name, a BookmarkStart, and a BookmarkEnd node.
    // Any whitespace in the names of bookmarks will be converted to underscores if we open the saved document with Microsoft Word. 
    // If we highlight the bookmark's name in Microsoft Word via Insert -> Links -> Bookmark, and press "Go To",
    // the cursor will jump to the text enclosed between the BookmarkStart and BookmarkEnd nodes.
    builder->StartBookmark(u"My Bookmark");
    builder->Write(u"Contents of MyBookmark.");
    builder->EndBookmark(u"My Bookmark");
    
    // Bookmarks are stored in this collection.
    ASSERT_EQ(u"My Bookmark", doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Name());
    
    doc->Save(get_ArtifactsDir() + u"Bookmarks.Insert.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Bookmarks.Insert.docx");
    
    ASSERT_EQ(u"My Bookmark", doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Name());
}

namespace gtest_test
{

TEST_F(ExBookmarks, Insert)
{
    s_instance->Insert();
}

} // namespace gtest_test

void ExBookmarks::CreateUpdateAndPrintBookmarks()
{
    // Create a document with three bookmarks, then use a custom document visitor implementation to print their contents.
    System::SharedPtr<Aspose::Words::Document> doc = CreateDocumentWithBookmarks(3);
    System::SharedPtr<Aspose::Words::BookmarkCollection> bookmarks = doc->get_Range()->get_Bookmarks();
    ASSERT_EQ(3, bookmarks->get_Count());
    //ExSkip
    
    PrintAllBookmarkInfo(bookmarks);
    
    // Bookmarks can be accessed in the bookmark collection by index or name, and their names can be updated.
    bookmarks->idx_get(0)->set_Name(System::String::Format(u"{0}_NewName", bookmarks->idx_get(0)->get_Name()));
    bookmarks->idx_get(u"MyBookmark_2")->set_Text(System::String::Format(u"Updated text contents of {0}", bookmarks->idx_get(1)->get_Name()));
    
    // Print all bookmarks again to see updated values.
    PrintAllBookmarkInfo(bookmarks);
}

namespace gtest_test
{

TEST_F(ExBookmarks, CreateUpdateAndPrintBookmarks)
{
    s_instance->CreateUpdateAndPrintBookmarks();
}

} // namespace gtest_test

void ExBookmarks::TableColumnBookmarks()
{
    //ExStart
    //ExFor:Bookmark.IsColumn
    //ExFor:Bookmark.FirstColumn
    //ExFor:Bookmark.LastColumn
    //ExSummary:Shows how to get information about table column bookmarks.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Table column bookmarks.doc");
    
    for (auto&& bookmark : System::IterateOver(doc->get_Range()->get_Bookmarks()))
    {
        // If a bookmark encloses columns of a table, it is a table column bookmark, and its IsColumn flag set to true.
        std::cout << System::String::Format(u"Bookmark: {0}{1}", bookmark->get_Name(), (bookmark->get_IsColumn() ? System::String(u" (Column)") : System::String(u""))) << std::endl;
        if (bookmark->get_IsColumn())
        {
            auto row = System::AsCast<Aspose::Words::Tables::Row>(bookmark->get_BookmarkStart()->GetAncestor(Aspose::Words::NodeType::Row));
            if (row != nullptr && bookmark->get_FirstColumn() < row->get_Cells()->get_Count())
            {
                // Print the contents of the first and last columns enclosed by the bookmark.
                std::cout << row->get_Cells()->idx_get(bookmark->get_FirstColumn())->GetText().TrimEnd(System::MakeArray<char16_t>({Aspose::Words::ControlChar::CellChar})) << std::endl;
                std::cout << row->get_Cells()->idx_get(bookmark->get_LastColumn())->GetText().TrimEnd(System::MakeArray<char16_t>({Aspose::Words::ControlChar::CellChar})) << std::endl;
            }
        }
    }
    //ExEnd
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    
    System::SharedPtr<Aspose::Words::Bookmark> firstTableColumnBookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"FirstTableColumnBookmark");
    System::SharedPtr<Aspose::Words::Bookmark> secondTableColumnBookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"SecondTableColumnBookmark");
    
    ASSERT_TRUE(firstTableColumnBookmark->get_IsColumn());
    ASSERT_EQ(1, firstTableColumnBookmark->get_FirstColumn());
    ASSERT_EQ(3, firstTableColumnBookmark->get_LastColumn());
    
    ASSERT_TRUE(secondTableColumnBookmark->get_IsColumn());
    ASSERT_EQ(0, secondTableColumnBookmark->get_FirstColumn());
    ASSERT_EQ(3, secondTableColumnBookmark->get_LastColumn());
}

namespace gtest_test
{

TEST_F(ExBookmarks, TableColumnBookmarks)
{
    s_instance->TableColumnBookmarks();
}

} // namespace gtest_test

void ExBookmarks::Remove()
{
    //ExStart
    //ExFor:BookmarkCollection.Clear
    //ExFor:BookmarkCollection.Count
    //ExFor:BookmarkCollection.Remove(Bookmark)
    //ExFor:BookmarkCollection.Remove(String)
    //ExFor:BookmarkCollection.RemoveAt
    //ExFor:Bookmark.Remove
    //ExSummary:Shows how to remove bookmarks from a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert five bookmarks with text inside their boundaries.
    for (int32_t i = 1; i <= 5; i++)
    {
        System::String bookmarkName = System::String(u"MyBookmark_") + i;
        
        builder->StartBookmark(bookmarkName);
        builder->Write(System::String::Format(u"Text inside {0}.", bookmarkName));
        builder->EndBookmark(bookmarkName);
        builder->InsertBreak(Aspose::Words::BreakType::ParagraphBreak);
    }
    
    // This collection stores bookmarks.
    System::SharedPtr<Aspose::Words::BookmarkCollection> bookmarks = doc->get_Range()->get_Bookmarks();
    
    ASSERT_EQ(5, bookmarks->get_Count());
    
    // There are several ways of removing bookmarks.
    // 1 -  Calling the bookmark's Remove method:
    bookmarks->idx_get(u"MyBookmark_1")->Remove();
    
    ASSERT_FALSE(bookmarks->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Bookmark>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Bookmark> b)>>([](System::SharedPtr<Aspose::Words::Bookmark> b) -> bool
    {
        return b->get_Name() == u"MyBookmark_1";
    }))));
    
    // 2 -  Passing the bookmark to the collection's Remove method:
    System::SharedPtr<Aspose::Words::Bookmark> bookmark = doc->get_Range()->get_Bookmarks()->idx_get(0);
    doc->get_Range()->get_Bookmarks()->Remove(bookmark);
    
    ASSERT_FALSE(bookmarks->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Bookmark>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Bookmark> b)>>([](System::SharedPtr<Aspose::Words::Bookmark> b) -> bool
    {
        return b->get_Name() == u"MyBookmark_2";
    }))));
    
    // 3 -  Removing a bookmark from the collection by name:
    doc->get_Range()->get_Bookmarks()->Remove(u"MyBookmark_3");
    
    ASSERT_FALSE(bookmarks->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Bookmark>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Bookmark> b)>>([](System::SharedPtr<Aspose::Words::Bookmark> b) -> bool
    {
        return b->get_Name() == u"MyBookmark_3";
    }))));
    
    // 4 -  Removing a bookmark at an index in the bookmark collection:
    doc->get_Range()->get_Bookmarks()->RemoveAt(0);
    
    ASSERT_FALSE(bookmarks->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Bookmark>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Bookmark> b)>>([](System::SharedPtr<Aspose::Words::Bookmark> b) -> bool
    {
        return b->get_Name() == u"MyBookmark_4";
    }))));
    
    // We can clear the entire bookmark collection.
    bookmarks->Clear();
    
    // The text that was inside the bookmarks is still present in the document.
    ASSERT_EQ(0, bookmarks->get_Count());
    ASSERT_EQ(System::String(u"Text inside MyBookmark_1.\r") + u"Text inside MyBookmark_2.\r" + u"Text inside MyBookmark_3.\r" + u"Text inside MyBookmark_4.\r" + u"Text inside MyBookmark_5.", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExBookmarks, Remove)
{
    s_instance->Remove();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
