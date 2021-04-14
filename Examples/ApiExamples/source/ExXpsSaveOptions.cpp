// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExXpsSaveOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace ApiExamples { namespace gtest_test {

class ExXpsSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExXpsSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExXpsSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExXpsSaveOptions> ExXpsSaveOptions::s_instance;

TEST_F(ExXpsSaveOptions, OutlineLevels)
{
    s_instance->OutlineLevels();
}

using ExXpsSaveOptions_BookFold_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExXpsSaveOptions::BookFold)>::type;

struct ExXpsSaveOptions_BookFold : public ExXpsSaveOptions,
                                   public ApiExamples::ExXpsSaveOptions,
                                   public ::testing::WithParamInterface<ExXpsSaveOptions_BookFold_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExXpsSaveOptions_BookFold, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->BookFold(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExXpsSaveOptions_BookFold, ::testing::ValuesIn(ExXpsSaveOptions_BookFold::TestCases()));

using ExXpsSaveOptions_OptimizeOutput_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExXpsSaveOptions::OptimizeOutput)>::type;

struct ExXpsSaveOptions_OptimizeOutput : public ExXpsSaveOptions,
                                         public ApiExamples::ExXpsSaveOptions,
                                         public ::testing::WithParamInterface<ExXpsSaveOptions_OptimizeOutput_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExXpsSaveOptions_OptimizeOutput, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->OptimizeOutput(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExXpsSaveOptions_OptimizeOutput, ::testing::ValuesIn(ExXpsSaveOptions_OptimizeOutput::TestCases()));

TEST_F(ExXpsSaveOptions, ExportExactPages)
{
    s_instance->ExportExactPages();
}

}} // namespace ApiExamples::gtest_test
