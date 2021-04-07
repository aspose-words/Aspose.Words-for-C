// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRange.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Replacing;
namespace ApiExamples { namespace gtest_test {

class ExRange : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExRange> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExRange>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExRange> ExRange::s_instance;

TEST_F(ExRange, Replace)
{
    s_instance->Replace();
}

using ExRange_ReplaceMatchCase_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRange::ReplaceMatchCase)>::type;

struct ExRange_ReplaceMatchCase : public ExRange, public ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_ReplaceMatchCase_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExRange_ReplaceMatchCase, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ReplaceMatchCase(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_ReplaceMatchCase, ::testing::ValuesIn(ExRange_ReplaceMatchCase::TestCases()));

using ExRange_ReplaceFindWholeWordsOnly_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRange::ReplaceFindWholeWordsOnly)>::type;

struct ExRange_ReplaceFindWholeWordsOnly : public ExRange,
                                           public ApiExamples::ExRange,
                                           public ::testing::WithParamInterface<ExRange_ReplaceFindWholeWordsOnly_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExRange_ReplaceFindWholeWordsOnly, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ReplaceFindWholeWordsOnly(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_ReplaceFindWholeWordsOnly, ::testing::ValuesIn(ExRange_ReplaceFindWholeWordsOnly::TestCases()));

using ExRange_IgnoreDeleted_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRange::IgnoreDeleted)>::type;

struct ExRange_IgnoreDeleted : public ExRange, public ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_IgnoreDeleted_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExRange_IgnoreDeleted, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreDeleted(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_IgnoreDeleted, ::testing::ValuesIn(ExRange_IgnoreDeleted::TestCases()));

using ExRange_IgnoreInserted_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRange::IgnoreInserted)>::type;

struct ExRange_IgnoreInserted : public ExRange, public ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_IgnoreInserted_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExRange_IgnoreInserted, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreInserted(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_IgnoreInserted, ::testing::ValuesIn(ExRange_IgnoreInserted::TestCases()));

using ExRange_IgnoreFields_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRange::IgnoreFields)>::type;

struct ExRange_IgnoreFields : public ExRange, public ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_IgnoreFields_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExRange_IgnoreFields, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreFields(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_IgnoreFields, ::testing::ValuesIn(ExRange_IgnoreFields::TestCases()));

TEST_F(ExRange, UpdateFieldsInRange)
{
    s_instance->UpdateFieldsInRange();
}

TEST_F(ExRange, ReplaceWithString)
{
    s_instance->ReplaceWithString();
}

TEST_F(ExRange, ReplaceWithRegex)
{
    s_instance->ReplaceWithRegex();
}

TEST_F(ExRange, ReplaceWithCallback)
{
    s_instance->ReplaceWithCallback();
}

TEST_F(ExRange, ConvertNumbersToHexadecimal)
{
    s_instance->ConvertNumbersToHexadecimal();
}

TEST_F(ExRange, ApplyParagraphFormat)
{
    s_instance->ApplyParagraphFormat();
}

TEST_F(ExRange, DeleteSelection)
{
    s_instance->DeleteSelection();
}

TEST_F(ExRange, RangesGetText)
{
    s_instance->RangesGetText();
}

using ExRange_UseLegacyOrder_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRange::UseLegacyOrder)>::type;

struct ExRange_UseLegacyOrder : public ExRange, public ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_UseLegacyOrder_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExRange_UseLegacyOrder, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UseLegacyOrder(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_UseLegacyOrder, ::testing::ValuesIn(ExRange_UseLegacyOrder::TestCases()));

using ExRange_UseSubstitutions_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRange::UseSubstitutions)>::type;

struct ExRange_UseSubstitutions : public ExRange, public ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_UseSubstitutions_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExRange_UseSubstitutions, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UseSubstitutions(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_UseSubstitutions, ::testing::ValuesIn(ExRange_UseSubstitutions::TestCases()));

TEST_F(ExRange, InsertDocumentAtReplace)
{
    s_instance->InsertDocumentAtReplace();
}

using ExRange_Direction_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRange::Direction)>::type;

struct ExRange_Direction : public ExRange, public ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_Direction_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(FindReplaceDirection::Backward),
            std::make_tuple(FindReplaceDirection::Forward),
        };
    }
};

TEST_P(ExRange_Direction, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Direction(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_Direction, ::testing::ValuesIn(ExRange_Direction::TestCases()));

}} // namespace ApiExamples::gtest_test
