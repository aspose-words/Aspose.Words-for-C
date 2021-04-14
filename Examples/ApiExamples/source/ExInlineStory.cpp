// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExInlineStory.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;
namespace ApiExamples { namespace gtest_test {

class ExInlineStory : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExInlineStory> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
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

using ExInlineStory_PositionFootnote_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExInlineStory::PositionFootnote)>::type;

struct ExInlineStory_PositionFootnote : public ExInlineStory,
                                        public ApiExamples::ExInlineStory,
                                        public ::testing::WithParamInterface<ExInlineStory_PositionFootnote_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(FootnotePosition::BeneathText),
            std::make_tuple(FootnotePosition::BottomOfPage),
        };
    }
};

TEST_P(ExInlineStory_PositionFootnote, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PositionFootnote(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExInlineStory_PositionFootnote, ::testing::ValuesIn(ExInlineStory_PositionFootnote::TestCases()));

using ExInlineStory_PositionEndnote_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExInlineStory::PositionEndnote)>::type;

struct ExInlineStory_PositionEndnote : public ExInlineStory,
                                       public ApiExamples::ExInlineStory,
                                       public ::testing::WithParamInterface<ExInlineStory_PositionEndnote_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(EndnotePosition::EndOfDocument),
            std::make_tuple(EndnotePosition::EndOfSection),
        };
    }
};

TEST_P(ExInlineStory_PositionEndnote, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PositionEndnote(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExInlineStory_PositionEndnote, ::testing::ValuesIn(ExInlineStory_PositionEndnote::TestCases()));

TEST_F(ExInlineStory, RefMarkNumberStyle)
{
    s_instance->RefMarkNumberStyle();
}

TEST_F(ExInlineStory, NumberingRule)
{
    s_instance->NumberingRule();
}

TEST_F(ExInlineStory, StartNumber)
{
    s_instance->StartNumber();
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
