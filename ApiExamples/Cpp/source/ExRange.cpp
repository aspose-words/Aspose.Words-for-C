// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRange.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <system/collections/list.h>
#include <system/string.h>

using namespace Aspose::Words;
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

TEST_F(ExRange, ReplaceSimple)
{
    s_instance->ReplaceSimple();
}

using IgnoreDeleted_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRange::IgnoreDeleted)>::type;

struct IgnoreDeletedVP : public ExRange, public ApiExamples::ExRange, public ::testing::WithParamInterface<IgnoreDeleted_Args>
{
    static std::vector<IgnoreDeleted_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(IgnoreDeletedVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreDeleted(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExRange, IgnoreDeletedVP, ::testing::ValuesIn(IgnoreDeletedVP::TestCases()));

using IgnoreInserted_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRange::IgnoreInserted)>::type;

struct IgnoreInsertedVP : public ExRange, public ApiExamples::ExRange, public ::testing::WithParamInterface<IgnoreInserted_Args>
{
    static std::vector<IgnoreInserted_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(IgnoreInsertedVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreInserted(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExRange, IgnoreInsertedVP, ::testing::ValuesIn(IgnoreInsertedVP::TestCases()));

using IgnoreFields_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRange::IgnoreFields)>::type;

struct IgnoreFieldsVP : public ExRange, public ApiExamples::ExRange, public ::testing::WithParamInterface<IgnoreFields_Args>
{
    static std::vector<IgnoreFields_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(IgnoreFieldsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreFields(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExRange, IgnoreFieldsVP, ::testing::ValuesIn(IgnoreFieldsVP::TestCases()));

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

TEST_F(ExRange, ReplaceWithInsertHtml)
{
    s_instance->ReplaceWithInsertHtml();
}

TEST_F(ExRange, ReplaceNumbersAsHex)
{
    s_instance->ReplaceNumbersAsHex();
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

using UseLegacyOrder_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRange::UseLegacyOrder)>::type;

struct UseLegacyOrderVP : public ExRange, public ApiExamples::ExRange, public ::testing::WithParamInterface<UseLegacyOrder_Args>
{
    static std::vector<UseLegacyOrder_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(UseLegacyOrderVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UseLegacyOrder(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExRange, UseLegacyOrderVP, ::testing::ValuesIn(UseLegacyOrderVP::TestCases()));

TEST_F(ExRange, UseSubstitutions)
{
    s_instance->UseSubstitutions();
}

TEST_F(ExRange, InsertDocumentAtReplace)
{
    s_instance->InsertDocumentAtReplace();
}

}} // namespace ApiExamples::gtest_test
