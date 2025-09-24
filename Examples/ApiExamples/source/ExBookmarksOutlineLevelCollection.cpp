// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBookmarksOutlineLevelCollection.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
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


using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3803395245u, ::Aspose::Words::ApiExamples::ExBookmarksOutlineLevelCollection, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExBookmarksOutlineLevelCollection : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExBookmarksOutlineLevelCollection> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExBookmarksOutlineLevelCollection>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExBookmarksOutlineLevelCollection> ExBookmarksOutlineLevelCollection::s_instance;

} // namespace gtest_test

void ExBookmarksOutlineLevelCollection::BookmarkLevels()
{
    //ExStart
    //ExFor:BookmarksOutlineLevelCollection
    //ExFor:BookmarksOutlineLevelCollection.Add(String, Int32)
    //ExFor:BookmarksOutlineLevelCollection.Clear
    //ExFor:BookmarksOutlineLevelCollection.Contains(String)
    //ExFor:BookmarksOutlineLevelCollection.Count
    //ExFor:BookmarksOutlineLevelCollection.IndexOfKey(String)
    //ExFor:BookmarksOutlineLevelCollection.Item(Int32)
    //ExFor:BookmarksOutlineLevelCollection.Item(String)
    //ExFor:BookmarksOutlineLevelCollection.Remove(String)
    //ExFor:BookmarksOutlineLevelCollection.RemoveAt(Int32)
    //ExFor:OutlineOptions.BookmarksOutlineLevels
    //ExSummary:Shows how to set outline levels for bookmarks.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a bookmark with another bookmark nested inside it.
    builder->StartBookmark(u"Bookmark 1");
    builder->Writeln(u"Text inside Bookmark 1.");
    
    builder->StartBookmark(u"Bookmark 2");
    builder->Writeln(u"Text inside Bookmark 1 and 2.");
    builder->EndBookmark(u"Bookmark 2");
    
    builder->Writeln(u"Text inside Bookmark 1.");
    builder->EndBookmark(u"Bookmark 1");
    
    // Insert another bookmark.
    builder->StartBookmark(u"Bookmark 3");
    builder->Writeln(u"Text inside Bookmark 3.");
    builder->EndBookmark(u"Bookmark 3");
    
    // When saving to .pdf, bookmarks can be accessed via a drop-down menu and used as anchors by most readers.
    // Bookmarks can also have numeric values for outline levels,
    // enabling lower level outline entries to hide higher-level child entries when collapsed in the reader.
    auto pdfSaveOptions = System::MakeObject<Aspose::Words::Saving::PdfSaveOptions>();
    System::SharedPtr<Aspose::Words::Saving::BookmarksOutlineLevelCollection> outlineLevels = pdfSaveOptions->get_OutlineOptions()->get_BookmarksOutlineLevels();
    
    outlineLevels->Add(u"Bookmark 1", 1);
    outlineLevels->Add(u"Bookmark 2", 2);
    outlineLevels->Add(u"Bookmark 3", 3);
    
    ASSERT_EQ(3, outlineLevels->get_Count());
    ASSERT_TRUE(outlineLevels->Contains(u"Bookmark 1"));
    ASSERT_EQ(1, outlineLevels->idx_get(0));
    ASSERT_EQ(2, outlineLevels->idx_get(u"Bookmark 2"));
    ASSERT_EQ(2, outlineLevels->IndexOfKey(u"Bookmark 3"));
    
    // We can remove two elements so that only the outline level designation for "Bookmark 1" is left.
    outlineLevels->RemoveAt(2);
    outlineLevels->Remove(u"Bookmark 2");
    
    // There are nine outline levels. Their numbering will be optimized during the save operation.
    // In this case, levels "5" and "9" will become "2" and "3".
    outlineLevels->Add(u"Bookmark 2", 5);
    outlineLevels->Add(u"Bookmark 3", 9);
    
    doc->Save(get_ArtifactsDir() + u"BookmarksOutlineLevelCollection.BookmarkLevels.pdf", pdfSaveOptions);
    
    // Emptying this collection will preserve the bookmarks and put them all on the same outline level.
    outlineLevels->Clear();
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExBookmarksOutlineLevelCollection, BookmarkLevels)
{
    s_instance->BookmarkLevels();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
