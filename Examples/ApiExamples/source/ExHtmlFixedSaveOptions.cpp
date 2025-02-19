// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHtmlFixedSaveOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples { namespace gtest_test {

class ExHtmlFixedSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExHtmlFixedSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExHtmlFixedSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExHtmlFixedSaveOptions> ExHtmlFixedSaveOptions::s_instance;

TEST_F(ExHtmlFixedSaveOptions, UseEncoding)
{
    s_instance->UseEncoding();
}

TEST_F(ExHtmlFixedSaveOptions, GetEncoding)
{
    s_instance->GetEncoding();
}

using ExHtmlFixedSaveOptions_ExportEmbeddedCss_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportEmbeddedCss)>::type;

struct ExHtmlFixedSaveOptions_ExportEmbeddedCss : public ExHtmlFixedSaveOptions,
                                                  public ApiExamples::ExHtmlFixedSaveOptions,
                                                  public ::testing::WithParamInterface<ExHtmlFixedSaveOptions_ExportEmbeddedCss_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExHtmlFixedSaveOptions_ExportEmbeddedCss, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportEmbeddedCss(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlFixedSaveOptions_ExportEmbeddedCss, ::testing::ValuesIn(ExHtmlFixedSaveOptions_ExportEmbeddedCss::TestCases()));

using ExHtmlFixedSaveOptions_ExportEmbeddedFonts_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportEmbeddedFonts)>::type;

struct ExHtmlFixedSaveOptions_ExportEmbeddedFonts : public ExHtmlFixedSaveOptions,
                                                    public ApiExamples::ExHtmlFixedSaveOptions,
                                                    public ::testing::WithParamInterface<ExHtmlFixedSaveOptions_ExportEmbeddedFonts_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExHtmlFixedSaveOptions_ExportEmbeddedFonts, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportEmbeddedFonts(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlFixedSaveOptions_ExportEmbeddedFonts, ::testing::ValuesIn(ExHtmlFixedSaveOptions_ExportEmbeddedFonts::TestCases()));

using ExHtmlFixedSaveOptions_ExportEmbeddedImages_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportEmbeddedImages)>::type;

struct ExHtmlFixedSaveOptions_ExportEmbeddedImages : public ExHtmlFixedSaveOptions,
                                                     public ApiExamples::ExHtmlFixedSaveOptions,
                                                     public ::testing::WithParamInterface<ExHtmlFixedSaveOptions_ExportEmbeddedImages_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExHtmlFixedSaveOptions_ExportEmbeddedImages, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportEmbeddedImages(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlFixedSaveOptions_ExportEmbeddedImages, ::testing::ValuesIn(ExHtmlFixedSaveOptions_ExportEmbeddedImages::TestCases()));

using ExHtmlFixedSaveOptions_ExportEmbeddedSvgs_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportEmbeddedSvgs)>::type;

struct ExHtmlFixedSaveOptions_ExportEmbeddedSvgs : public ExHtmlFixedSaveOptions,
                                                   public ApiExamples::ExHtmlFixedSaveOptions,
                                                   public ::testing::WithParamInterface<ExHtmlFixedSaveOptions_ExportEmbeddedSvgs_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExHtmlFixedSaveOptions_ExportEmbeddedSvgs, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportEmbeddedSvgs(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlFixedSaveOptions_ExportEmbeddedSvgs, ::testing::ValuesIn(ExHtmlFixedSaveOptions_ExportEmbeddedSvgs::TestCases()));

using ExHtmlFixedSaveOptions_ExportFormFields_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportFormFields)>::type;

struct ExHtmlFixedSaveOptions_ExportFormFields : public ExHtmlFixedSaveOptions,
                                                 public ApiExamples::ExHtmlFixedSaveOptions,
                                                 public ::testing::WithParamInterface<ExHtmlFixedSaveOptions_ExportFormFields_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExHtmlFixedSaveOptions_ExportFormFields, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportFormFields(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlFixedSaveOptions_ExportFormFields, ::testing::ValuesIn(ExHtmlFixedSaveOptions_ExportFormFields::TestCases()));

using ExHtmlFixedSaveOptions_HorizontalAlignment_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::HorizontalAlignment)>::type;

struct ExHtmlFixedSaveOptions_HorizontalAlignment : public ExHtmlFixedSaveOptions,
                                                    public ApiExamples::ExHtmlFixedSaveOptions,
                                                    public ::testing::WithParamInterface<ExHtmlFixedSaveOptions_HorizontalAlignment_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(HtmlFixedPageHorizontalAlignment::Center),
            std::make_tuple(HtmlFixedPageHorizontalAlignment::Left),
            std::make_tuple(HtmlFixedPageHorizontalAlignment::Right),
        };
    }
};

TEST_P(ExHtmlFixedSaveOptions_HorizontalAlignment, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->HorizontalAlignment(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlFixedSaveOptions_HorizontalAlignment, ::testing::ValuesIn(ExHtmlFixedSaveOptions_HorizontalAlignment::TestCases()));

TEST_F(ExHtmlFixedSaveOptions, PageMargins)
{
    s_instance->PageMargins();
}

TEST_F(ExHtmlFixedSaveOptions, PageMarginsException)
{
    s_instance->PageMarginsException();
}

using ExHtmlFixedSaveOptions_UsingMachineFonts_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::UsingMachineFonts)>::type;

struct ExHtmlFixedSaveOptions_UsingMachineFonts : public ExHtmlFixedSaveOptions,
                                                  public ApiExamples::ExHtmlFixedSaveOptions,
                                                  public ::testing::WithParamInterface<ExHtmlFixedSaveOptions_UsingMachineFonts_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlFixedSaveOptions_UsingMachineFonts, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UsingMachineFonts(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlFixedSaveOptions_UsingMachineFonts, ::testing::ValuesIn(ExHtmlFixedSaveOptions_UsingMachineFonts::TestCases()));

TEST_F(ExHtmlFixedSaveOptions, ResourceSavingCallback)
{
    s_instance->ResourceSavingCallback();
}

TEST_F(ExHtmlFixedSaveOptions, HtmlFixedResourceFolder)
{
    s_instance->HtmlFixedResourceFolder();
}

}} // namespace ApiExamples::gtest_test
