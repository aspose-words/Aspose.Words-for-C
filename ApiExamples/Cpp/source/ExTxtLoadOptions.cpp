// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTxtLoadOptions.h"

#include <Aspose.Words.Cpp/Model/Document/TxtLeadingSpacesOptions.h>
#include <Aspose.Words.Cpp/Model/Document/TxtTrailingSpacesOptions.h>

using namespace Aspose::Words;
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

using DetectNumberingWithWhitespaces_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtLoadOptions::DetectNumberingWithWhitespaces)>::type;

struct DetectNumberingWithWhitespacesVP : public ExTxtLoadOptions,
                                          public ApiExamples::ExTxtLoadOptions,
                                          public ::testing::WithParamInterface<DetectNumberingWithWhitespaces_Args>
{
    static std::vector<DetectNumberingWithWhitespaces_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(DetectNumberingWithWhitespacesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DetectNumberingWithWhitespaces(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtLoadOptions, DetectNumberingWithWhitespacesVP, ::testing::ValuesIn(DetectNumberingWithWhitespacesVP::TestCases()));

using TrailSpaces_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtLoadOptions::TrailSpaces)>::type;

struct TrailSpacesVP : public ExTxtLoadOptions, public ApiExamples::ExTxtLoadOptions, public ::testing::WithParamInterface<TrailSpaces_Args>
{
    static std::vector<TrailSpaces_Args> TestCases()
    {
        return {
            std::make_tuple(TxtLeadingSpacesOptions::Preserve, TxtTrailingSpacesOptions::Preserve),
            std::make_tuple(TxtLeadingSpacesOptions::ConvertToIndent, TxtTrailingSpacesOptions::Preserve),
            std::make_tuple(TxtLeadingSpacesOptions::Trim, TxtTrailingSpacesOptions::Trim),
        };
    }
};

TEST_P(TrailSpacesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TrailSpaces(get<0>(params), get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtLoadOptions, TrailSpacesVP, ::testing::ValuesIn(TrailSpacesVP::TestCases()));

TEST_F(ExTxtLoadOptions, DetectDocumentDirection)
{
    s_instance->DetectDocumentDirection();
}

}} // namespace ApiExamples::gtest_test
