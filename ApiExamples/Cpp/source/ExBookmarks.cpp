// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBookmarks.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>

using namespace Aspose::Words;
namespace ApiExamples { namespace gtest_test {

class ExBookmarks : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExBookmarks> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExBookmarks>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExBookmarks> ExBookmarks::s_instance;

TEST_F(ExBookmarks, Insert)
{
    s_instance->Insert();
}

TEST_F(ExBookmarks, CreateUpdateAndPrintBookmarks)
{
    s_instance->CreateUpdateAndPrintBookmarks();
}

TEST_F(ExBookmarks, TableColumnBookmarks)
{
    s_instance->TableColumnBookmarks();
}

TEST_F(ExBookmarks, Remove)
{
    s_instance->Remove();
}

}} // namespace ApiExamples::gtest_test
