// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBookmarks.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/console.h>
#include <system/collections/ienumerator.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1733492527u, ::ApiExamples::ExBookmarks::BookmarkInfoPrinter, ThisTypeBaseTypesInfo);

Aspose::Words::VisitorAction ExBookmarks::BookmarkInfoPrinter::VisitBookmarkStart(SharedPtr<Aspose::Words::BookmarkStart> bookmarkStart)
{
    System::Console::WriteLine(u"BookmarkStart name: \"{0}\", Content: \"{1}\"", System::ObjectExt::Box<String>(bookmarkStart->get_Name()), System::ObjectExt::Box<String>(bookmarkStart->get_Bookmark()->get_Text()));
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExBookmarks::BookmarkInfoPrinter::VisitBookmarkEnd(SharedPtr<Aspose::Words::BookmarkEnd> bookmarkEnd)
{
    System::Console::WriteLine(u"BookmarkEnd name: \"{0}\"", System::ObjectExt::Box<String>(bookmarkEnd->get_Name()));
    return Aspose::Words::VisitorAction::Continue;
}

SharedPtr<Aspose::Words::Document> ExBookmarks::CreateDocumentWithBookmarks()
{
    auto builder = MakeObject<DocumentBuilder>();
    SharedPtr<Document> doc = builder->get_Document();

    // An empty document has just one empty paragraph by default
    SharedPtr<Paragraph> p = doc->get_FirstSection()->get_Body()->get_FirstParagraph();

    // Add several bookmarks to the document
    for (int i = 1; i <= 3; i++)
    {
        String bookmarkName = String(u"MyBookmark ") + i;

        p->AppendChild(MakeObject<Run>(doc, u"Text before bookmark."));

        p->AppendChild(MakeObject<BookmarkStart>(doc, bookmarkName));
        p->AppendChild(MakeObject<Run>(doc, String(u"Text content of ") + bookmarkName));
        p->AppendChild(MakeObject<BookmarkEnd>(doc, bookmarkName));

        p->AppendChild(MakeObject<Run>(doc, u"Text after bookmark.\r\n"));
    }

    return builder->get_Document();
}

void ExBookmarks::PrintAllBookmarkInfo(SharedPtr<Aspose::Words::BookmarkCollection> bookmarks)
{
    // Create a DocumentVisitor
    auto bookmarkVisitor = MakeObject<ExBookmarks::BookmarkInfoPrinter>();

    // Get the enumerator from the document's BookmarkCollection and iterate over the bookmarks
    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Bookmark>>> enumerator = bookmarks->GetEnumerator();
        while (enumerator->MoveNext())
        {
            SharedPtr<Bookmark> currentBookmark = enumerator->get_Current();

            // Accept our DocumentVisitor it to print information about our bookmarks
            if (currentBookmark != nullptr)
            {
                currentBookmark->get_BookmarkStart()->Accept(bookmarkVisitor);
                currentBookmark->get_BookmarkEnd()->Accept(bookmarkVisitor);

                // Prints a blank line
                System::Console::WriteLine(currentBookmark->get_BookmarkStart()->GetText());
            }
        }
    }
}

namespace gtest_test
{

class ExBookmarks : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExBookmarks> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExBookmarks>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExBookmarks> ExBookmarks::s_instance;

} // namespace gtest_test

void ExBookmarks::CreateUpdateAndPrintBookmarks()
{
    // Create a document with 3 bookmarks: "MyBookmark 1", "MyBookmark 2", "MyBookmark 3"
    SharedPtr<Document> doc = CreateDocumentWithBookmarks();
    SharedPtr<BookmarkCollection> bookmarks = doc->get_Range()->get_Bookmarks();
    ASSERT_EQ(3, bookmarks->get_Count());
    //ExSkip
    ASSERT_EQ(u"MyBookmark 1", bookmarks->idx_get(0)->get_Name());
    //ExSkip
    ASSERT_EQ(u"Text content of MyBookmark 2", bookmarks->idx_get(1)->get_Text());
    //ExSkip

    // Look at initial values of our bookmarks
    PrintAllBookmarkInfo(bookmarks);

    // Obtain bookmarks from a bookmark collection by index/name and update their values
    bookmarks->idx_get(0)->set_Name(String(u"Updated name of ") + bookmarks->idx_get(0)->get_Name());
    bookmarks->idx_get(u"MyBookmark 2")->set_Text(String(u"Updated text content of ") + bookmarks->idx_get(1)->get_Name());
    // Remove the latest bookmark
    // The bookmarked text is not deleted
    bookmarks->idx_get(2)->Remove();

    bookmarks = doc->get_Range()->get_Bookmarks();
    // Check that we have 2 bookmarks after the latest bookmark was deleted
    ASSERT_EQ(2, bookmarks->get_Count());
    ASSERT_EQ(u"Updated name of MyBookmark 1", bookmarks->idx_get(0)->get_Name());
    //ExSkip
    ASSERT_EQ(u"Updated text content of MyBookmark 2", bookmarks->idx_get(1)->get_Text());
    //ExSkip

    // Look at updated values of our bookmarks
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
    //ExSummary:Shows how to get information about table column bookmark.
    auto doc = MakeObject<Document>(MyDir + u"TableColumnBookmark.doc");
    for (auto bookmark : System::IterateOver(doc->get_Range()->get_Bookmarks()))
    {
        System::Console::WriteLine(u"Bookmark: {0}{1}", System::ObjectExt::Box<String>(bookmark->get_Name()), System::ObjectExt::Box<String>(bookmark->get_IsColumn() ? String(u" (Column)") : String(u"")));
        if (bookmark->get_IsColumn())
        {
            auto row = System::DynamicCast_noexcept<Aspose::Words::Tables::Row>(bookmark->get_BookmarkStart()->GetAncestor(Aspose::Words::NodeType::Row));
            if (row != nullptr && bookmark->get_FirstColumn() < row->get_Cells()->get_Count())
            {
                // Print text from the first and last cells containing in bookmark
                System::Console::WriteLine(row->get_Cells()->idx_get(bookmark->get_FirstColumn())->GetText().TrimEnd(MakeArray<char16_t>({ControlChar::CellChar})));
                System::Console::WriteLine(row->get_Cells()->idx_get(bookmark->get_LastColumn())->GetText().TrimEnd(MakeArray<char16_t>({ControlChar::CellChar})));
            }
        }
    }
    //ExEnd

    SharedPtr<Bookmark> firstTableColumnBookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"FirstTableColumnBookmark");
    SharedPtr<Bookmark> secondTableColumnBookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"SecondTableColumnBookmark");

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

void ExBookmarks::ClearBookmarks()
{
    //ExStart
    //ExFor:BookmarkCollection.Clear
    //ExSummary:Shows how to remove all bookmarks from a document.
    // Open a document with 3 bookmarks: "MyBookmark1", "My_Bookmark2", "MyBookmark3"
    auto doc = MakeObject<Document>(MyDir + u"Bookmarks.docx");

    // Remove all bookmarks from the document
    // The bookmarked text is not deleted
    doc->get_Range()->get_Bookmarks()->Clear();
    //ExEnd

    // Verify that the bookmarks were removed
    ASSERT_EQ(0, doc->get_Range()->get_Bookmarks()->get_Count());
}

namespace gtest_test
{

TEST_F(ExBookmarks, ClearBookmarks)
{
    s_instance->ClearBookmarks();
}

} // namespace gtest_test

void ExBookmarks::RemoveBookmarkFromBookmarkCollection()
{
    //ExStart
    //ExFor:BookmarkCollection.Remove(Bookmark)
    //ExFor:BookmarkCollection.Remove(String)
    //ExFor:BookmarkCollection.RemoveAt
    //ExSummary:Shows how to remove bookmarks from a document using different methods.
    // Open a document with 3 bookmarks: "MyBookmark1", "My_Bookmark2", "MyBookmark3"
    auto doc = MakeObject<Document>(MyDir + u"Bookmarks.docx");

    // Remove a particular bookmark from the document
    SharedPtr<Bookmark> bookmark = doc->get_Range()->get_Bookmarks()->idx_get(0);
    doc->get_Range()->get_Bookmarks()->Remove(bookmark);

    // Remove a bookmark by specified name
    doc->get_Range()->get_Bookmarks()->Remove(u"My_Bookmark2");

    // Remove a bookmark at the specified index
    doc->get_Range()->get_Bookmarks()->RemoveAt(0);
    //ExEnd

    // In docx we have additional hidden bookmark "_GoBack"
    // When we check bookmarks count, the result will be 1 instead of 0
    ASSERT_EQ(1, doc->get_Range()->get_Bookmarks()->get_Count());
}

namespace gtest_test
{

TEST_F(ExBookmarks, RemoveBookmarkFromBookmarkCollection)
{
    s_instance->RemoveBookmarkFromBookmarkCollection();
}

} // namespace gtest_test

void ExBookmarks::ReplaceBookmarkUnderscoresWithWhitespaces()
{
    //ExStart
    //ExFor:Bookmark.Name
    //ExSummary:Shows how to replace elements in bookmark name
    // Open a document with 3 bookmarks: "MyBookmark1", "My_Bookmark2", "MyBookmark3"
    auto doc = MakeObject<Document>(MyDir + u"Bookmarks.docx");
    ASSERT_EQ(u"MyBookmark3", doc->get_Range()->get_Bookmarks()->idx_get(2)->get_Name());
    //ExSkip

    // MS Word document does not support bookmark names with whitespaces by default
    // If you have document which contains bookmark names with underscores, you can simply replace them to whitespaces
    for (auto bookmark : System::IterateOver(doc->get_Range()->get_Bookmarks()))
    {
        bookmark->set_Name(bookmark->get_Name().Replace(u"_", u" "));
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExBookmarks, ReplaceBookmarkUnderscoresWithWhitespaces)
{
    s_instance->ReplaceBookmarkUnderscoresWithWhitespaces();
}

} // namespace gtest_test

} // namespace ApiExamples
