// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRtfSaveOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
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

using ExRtfSaveOptions_ExportImages_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRtfSaveOptions::ExportImages)>::type;

struct ExRtfSaveOptions_ExportImages : public ExRtfSaveOptions,
                                       public ApiExamples::ExRtfSaveOptions,
                                       public ::testing::WithParamInterface<ExRtfSaveOptions_ExportImages_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExRtfSaveOptions_ExportImages, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportImages(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRtfSaveOptions_ExportImages, ::testing::ValuesIn(ExRtfSaveOptions_ExportImages::TestCases()));

using ExRtfSaveOptions_SkipMono_SkipMono_SaveImagesAsWmf_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRtfSaveOptions::SaveImagesAsWmf)>::type;

struct ExRtfSaveOptions_SkipMono_SkipMono_SaveImagesAsWmf : public ExRtfSaveOptions,
                                                            public ApiExamples::ExRtfSaveOptions,
                                                            public ::testing::WithParamInterface<ExRtfSaveOptions_SkipMono_SkipMono_SaveImagesAsWmf_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExRtfSaveOptions_SkipMono_SkipMono_SaveImagesAsWmf, Test)
{
    RecordProperty("category", "SkipMono");
    RecordProperty("category1", "SkipMono");
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SaveImagesAsWmf(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRtfSaveOptions_SkipMono_SkipMono_SaveImagesAsWmf,
                         ::testing::ValuesIn(ExRtfSaveOptions_SkipMono_SkipMono_SaveImagesAsWmf::TestCases()));

}} // namespace ApiExamples::gtest_test
