// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTxtLoadOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;
namespace ApiExamples { namespace gtest_test {

class ExTxtLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExTxtLoadOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExTxtLoadOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExTxtLoadOptions> ExTxtLoadOptions::s_instance;

using ExTxtLoadOptions_DetectNumberingWithWhitespaces_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtLoadOptions::DetectNumberingWithWhitespaces)>::type;

struct ExTxtLoadOptions_DetectNumberingWithWhitespaces : public ExTxtLoadOptions,
                                                         public ApiExamples::ExTxtLoadOptions,
                                                         public ::testing::WithParamInterface<ExTxtLoadOptions_DetectNumberingWithWhitespaces_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExTxtLoadOptions_DetectNumberingWithWhitespaces, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DetectNumberingWithWhitespaces(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtLoadOptions_DetectNumberingWithWhitespaces, ::testing::ValuesIn(ExTxtLoadOptions_DetectNumberingWithWhitespaces::TestCases()));

using ExTxtLoadOptions_TrailSpaces_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtLoadOptions::TrailSpaces)>::type;

struct ExTxtLoadOptions_TrailSpaces : public ExTxtLoadOptions,
                                      public ApiExamples::ExTxtLoadOptions,
                                      public ::testing::WithParamInterface<ExTxtLoadOptions_TrailSpaces_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(TxtLeadingSpacesOptions::Preserve, TxtTrailingSpacesOptions::Preserve),
            std::make_tuple(TxtLeadingSpacesOptions::ConvertToIndent, TxtTrailingSpacesOptions::Preserve),
            std::make_tuple(TxtLeadingSpacesOptions::Trim, TxtTrailingSpacesOptions::Trim),
        };
    }
};

TEST_P(ExTxtLoadOptions_TrailSpaces, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TrailSpaces(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtLoadOptions_TrailSpaces, ::testing::ValuesIn(ExTxtLoadOptions_TrailSpaces::TestCases()));

TEST_F(ExTxtLoadOptions, DetectDocumentDirection)
{
    s_instance->DetectDocumentDirection();
}

}} // namespace ApiExamples::gtest_test
