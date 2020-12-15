// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRtfLoadOptions.h"

namespace ApiExamples { namespace gtest_test {

class ExRtfLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExRtfLoadOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExRtfLoadOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExRtfLoadOptions> ExRtfLoadOptions::s_instance;

using RecognizeUtf8Text_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRtfLoadOptions::RecognizeUtf8Text)>::type;

struct RecognizeUtf8TextVP : public ExRtfLoadOptions, public ApiExamples::ExRtfLoadOptions, public ::testing::WithParamInterface<RecognizeUtf8Text_Args>
{
    static std::vector<RecognizeUtf8Text_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(RecognizeUtf8TextVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RecognizeUtf8Text(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExRtfLoadOptions, RecognizeUtf8TextVP, ::testing::ValuesIn(RecognizeUtf8TextVP::TestCases()));

}} // namespace ApiExamples::gtest_test
