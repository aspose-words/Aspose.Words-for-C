// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRtfSaveOptions.h"

namespace ApiExamples { namespace gtest_test {

class ExRtfSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExRtfSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExRtfSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExRtfSaveOptions> ExRtfSaveOptions::s_instance;

using ExportImages_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRtfSaveOptions::ExportImages)>::type;

struct ExportImagesVP : public ExRtfSaveOptions, public ApiExamples::ExRtfSaveOptions, public ::testing::WithParamInterface<ExportImages_Args>
{
    static std::vector<ExportImages_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExportImagesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportImages(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExRtfSaveOptions, ExportImagesVP, ::testing::ValuesIn(ExportImagesVP::TestCases()));

TEST_F(ExRtfSaveOptions, SkipMono_SaveImagesAsWmf)
{
    RecordProperty("category", "SkipMono");

    s_instance->SaveImagesAsWmf();
}

}} // namespace ApiExamples::gtest_test
