// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExWordML2003SaveOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples { namespace gtest_test {

class ExWordML2003SaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExWordML2003SaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExWordML2003SaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExWordML2003SaveOptions> ExWordML2003SaveOptions::s_instance;

using ExWordML2003SaveOptions_PrettyFormat_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExWordML2003SaveOptions::PrettyFormat)>::type;

struct ExWordML2003SaveOptions_PrettyFormat : public ExWordML2003SaveOptions,
                                              public ApiExamples::ExWordML2003SaveOptions,
                                              public ::testing::WithParamInterface<ExWordML2003SaveOptions_PrettyFormat_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExWordML2003SaveOptions_PrettyFormat, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PrettyFormat(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExWordML2003SaveOptions_PrettyFormat, ::testing::ValuesIn(ExWordML2003SaveOptions_PrettyFormat::TestCases()));

using ExWordML2003SaveOptions_MemoryOptimization_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExWordML2003SaveOptions::MemoryOptimization)>::type;

struct ExWordML2003SaveOptions_MemoryOptimization : public ExWordML2003SaveOptions,
                                                    public ApiExamples::ExWordML2003SaveOptions,
                                                    public ::testing::WithParamInterface<ExWordML2003SaveOptions_MemoryOptimization_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExWordML2003SaveOptions_MemoryOptimization, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MemoryOptimization(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExWordML2003SaveOptions_MemoryOptimization, ::testing::ValuesIn(ExWordML2003SaveOptions_MemoryOptimization::TestCases()));

}} // namespace ApiExamples::gtest_test
