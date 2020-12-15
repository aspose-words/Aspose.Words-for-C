// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExXpsSaveOptions.h"

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

using OptimizeOutput_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExXpsSaveOptions::OptimizeOutput)>::type;

struct OptimizeOutputVP : public ExXpsSaveOptions, public ApiExamples::ExXpsSaveOptions, public ::testing::WithParamInterface<OptimizeOutput_Args>
{
    static std::vector<OptimizeOutput_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(OptimizeOutputVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->OptimizeOutput(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExXpsSaveOptions, OptimizeOutputVP, ::testing::ValuesIn(OptimizeOutputVP::TestCases()));

}} // namespace ApiExamples::gtest_test
