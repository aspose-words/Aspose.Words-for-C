// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExPsSaveOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace ApiExamples { namespace gtest_test {

class ExPsSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExPsSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExPsSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExPsSaveOptions> ExPsSaveOptions::s_instance;

using ExPsSaveOptions_UseBookFoldPrintingSettings_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExPsSaveOptions::UseBookFoldPrintingSettings)>::type;

struct ExPsSaveOptions_UseBookFoldPrintingSettings : public ExPsSaveOptions,
                                                     public ApiExamples::ExPsSaveOptions,
                                                     public ::testing::WithParamInterface<ExPsSaveOptions_UseBookFoldPrintingSettings_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPsSaveOptions_UseBookFoldPrintingSettings, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UseBookFoldPrintingSettings(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPsSaveOptions_UseBookFoldPrintingSettings, ::testing::ValuesIn(ExPsSaveOptions_UseBookFoldPrintingSettings::TestCases()));

}} // namespace ApiExamples::gtest_test
