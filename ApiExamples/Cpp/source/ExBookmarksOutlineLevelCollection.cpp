// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: Aspose.Pdf is not supported
#include "ExBookmarksOutlineLevelCollection.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OutlineOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/BookmarksOutlineLevelCollection.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples {

namespace gtest_test
{

class ExBookmarksOutlineLevelCollection : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExBookmarksOutlineLevelCollection> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExBookmarksOutlineLevelCollection>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExBookmarksOutlineLevelCollection> ExBookmarksOutlineLevelCollection::s_instance;

} // namespace gtest_test

void ExBookmarksOutlineLevelCollection::BookmarkLevels()
{
    //ExStart
    //ExFor:BookmarksOutlineLevelCollection
    //ExFor:BookmarksOutlineLevelCollection.Add(String, Int32)
    //ExFor:BookmarksOutlineLevelCollection.Clear
    //ExFor:BookmarksOutlineLevelCollection.Contains(System.String)
    //ExFor:BookmarksOutlineLevelCollection.Count
    //ExFor:BookmarksOutlineLevelCollection.IndexOfKey(System.String)
    //ExFor:BookmarksOutlineLevelCollection.Item(System.Int32)
    //ExFor:BookmarksOutlineLevelCollection.Item(System.String)
    //ExFor:BookmarksOutlineLevelCollection.Remove(System.String)
    //ExFor:BookmarksOutlineLevelCollection.RemoveAt(System.Int32)
    //ExFor:OutlineOptions.BookmarksOutlineLevels
    //ExSummary:Shows how to set outline levels for bookmarks.
    // Open a blank document, create a DocumentBuilder, and use the builder to add some text wrapped inside bookmarks
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Note that whitespaces in bookmark names will be converted into underscores when saved to Microsoft Word formats
    // such as .doc and .docx, but will be preserved in other formats like .pdf or .xps
    builder->StartBookmark(u"Bookmark 1");
    builder->Writeln(u"Text inside Bookmark 1.");

    builder->StartBookmark(u"Bookmark 2");
    builder->Writeln(u"Text inside Bookmark 1 and 2.");
    builder->EndBookmark(u"Bookmark 2");

    builder->Writeln(u"Text inside Bookmark 1.");
    builder->EndBookmark(u"Bookmark 1");

    builder->StartBookmark(u"Bookmark 3");
    builder->Writeln(u"Text inside Bookmark 3.");
    builder->EndBookmark(u"Bookmark 3");

    // We can specify outline levels for our bookmarks so that they show up in the table of contents and are indented by an amount
    // of space proportional to the indent level in a SaveOptions object
    // Some pdf/xps readers such as Google Chrome also allow the collapsing of all higher level bookmarks by adjacent lower level bookmarks
    // This feature applies to .pdf or .xps file formats, so only their respective SaveOptions subclasses will support it
    auto pdfSaveOptions = MakeObject<PdfSaveOptions>();
    SharedPtr<BookmarksOutlineLevelCollection> outlineLevels = pdfSaveOptions->get_OutlineOptions()->get_BookmarksOutlineLevels();

    outlineLevels->Add(u"Bookmark 1", 1);
    outlineLevels->Add(u"Bookmark 2", 2);
    outlineLevels->Add(u"Bookmark 3", 3);

    ASSERT_EQ(3, outlineLevels->get_Count());
    ASSERT_TRUE(outlineLevels->Contains(u"Bookmark 1"));
    ASSERT_EQ(1, outlineLevels->idx_get(0));
    ASSERT_EQ(2, outlineLevels->idx_get(u"Bookmark 2"));
    ASSERT_EQ(2, outlineLevels->IndexOfKey(u"Bookmark 3"));

    // We can remove two elements so that only the outline level designation for "Bookmark 1" is left
    outlineLevels->RemoveAt(2);
    outlineLevels->Remove(u"Bookmark 2");

    // We have 9 bookmark levels to work with, and bookmark levels are also sorted in ascending order,
    // and get numbered in succession along that order
    // Practically this means that our three levels "1, 5, 9", will be seen as "1, 2, 3" in the output
    outlineLevels->Add(u"Bookmark 2", 5);
    outlineLevels->Add(u"Bookmark 3", 9);

    // Save the document as a .pdf and find links to the bookmarks and their outline levels
    doc->Save(ArtifactsDir + u"BookmarksOutlineLevelCollection.BookmarkLevels.pdf", pdfSaveOptions);

    // We can empty this dictionary to remove the contents table
    outlineLevels->Clear();
    //ExEnd

    // CPPDEFECT: PdfBookmarkEditor is not supported
}

namespace gtest_test
{

TEST_F(ExBookmarksOutlineLevelCollection, BookmarkLevels)
{
    s_instance->BookmarkLevels();
}

} // namespace gtest_test

} // namespace ApiExamples
