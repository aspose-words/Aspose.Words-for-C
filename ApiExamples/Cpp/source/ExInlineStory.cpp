// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExInlineStory.h"

namespace ApiExamples { namespace gtest_test {

class ExInlineStory : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExInlineStory> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExInlineStory>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExInlineStory> ExInlineStory::s_instance;

TEST_F(ExInlineStory, Footnotes)
{
    s_instance->Footnotes();
}

TEST_F(ExInlineStory, Endnotes)
{
    s_instance->Endnotes();
}

TEST_F(ExInlineStory, AddFootnote)
{
    s_instance->AddFootnote();
}

TEST_F(ExInlineStory, FootnoteEndnote)
{
    s_instance->FootnoteEndnote();
}

TEST_F(ExInlineStory, AddComment)
{
    s_instance->AddComment();
}

TEST_F(ExInlineStory, InlineStoryRevisions)
{
    s_instance->InlineStoryRevisions();
}

TEST_F(ExInlineStory, InsertInlineStoryNodes)
{
    s_instance->InsertInlineStoryNodes();
}

TEST_F(ExInlineStory, DeleteShapes)
{
    s_instance->DeleteShapes();
}

}} // namespace ApiExamples::gtest_test
