// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHtmlFixedSaveOptions.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/HtmlFixedPageHorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Saving/ResourceSavingArgs.h>

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

using ExportEmbeddedCSS_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportEmbeddedCSS)>::type;

struct ExportEmbeddedCSSVP : public ExHtmlFixedSaveOptions,
                             public ApiExamples::ExHtmlFixedSaveOptions,
                             public ::testing::WithParamInterface<ExportEmbeddedCSS_Args>
{
    static std::vector<ExportEmbeddedCSS_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExportEmbeddedCSSVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportEmbeddedCSS(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlFixedSaveOptions, ExportEmbeddedCSSVP, ::testing::ValuesIn(ExportEmbeddedCSSVP::TestCases()));

using ExportEmbeddedFonts_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportEmbeddedFonts)>::type;

struct ExportEmbeddedFontsVP : public ExHtmlFixedSaveOptions,
                               public ApiExamples::ExHtmlFixedSaveOptions,
                               public ::testing::WithParamInterface<ExportEmbeddedFonts_Args>
{
    static std::vector<ExportEmbeddedFonts_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExportEmbeddedFontsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportEmbeddedFonts(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlFixedSaveOptions, ExportEmbeddedFontsVP, ::testing::ValuesIn(ExportEmbeddedFontsVP::TestCases()));

using ExportEmbeddedImages_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportEmbeddedImages)>::type;

struct ExportEmbeddedImagesVP : public ExHtmlFixedSaveOptions,
                                public ApiExamples::ExHtmlFixedSaveOptions,
                                public ::testing::WithParamInterface<ExportEmbeddedImages_Args>
{
    static std::vector<ExportEmbeddedImages_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExportEmbeddedImagesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportEmbeddedImages(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlFixedSaveOptions, ExportEmbeddedImagesVP, ::testing::ValuesIn(ExportEmbeddedImagesVP::TestCases()));

using ExportEmbeddedSvgs_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportEmbeddedSvgs)>::type;

struct ExportEmbeddedSvgsVP : public ExHtmlFixedSaveOptions,
                              public ApiExamples::ExHtmlFixedSaveOptions,
                              public ::testing::WithParamInterface<ExportEmbeddedSvgs_Args>
{
    static std::vector<ExportEmbeddedSvgs_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExportEmbeddedSvgsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportEmbeddedSvgs(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlFixedSaveOptions, ExportEmbeddedSvgsVP, ::testing::ValuesIn(ExportEmbeddedSvgsVP::TestCases()));

using ExportFormFields_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportFormFields)>::type;

struct ExportFormFieldsVP : public ExHtmlFixedSaveOptions,
                            public ApiExamples::ExHtmlFixedSaveOptions,
                            public ::testing::WithParamInterface<ExportFormFields_Args>
{
    static std::vector<ExportFormFields_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExportFormFieldsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportFormFields(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlFixedSaveOptions, ExportFormFieldsVP, ::testing::ValuesIn(ExportFormFieldsVP::TestCases()));

TEST_F(ExHtmlFixedSaveOptions, AddCssClassNamesPrefix)
{
    s_instance->AddCssClassNamesPrefix();
}

using HorizontalAlignment_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::HorizontalAlignment)>::type;

struct HorizontalAlignmentVP : public ExHtmlFixedSaveOptions,
                               public ApiExamples::ExHtmlFixedSaveOptions,
                               public ::testing::WithParamInterface<HorizontalAlignment_Args>
{
    static std::vector<HorizontalAlignment_Args> TestCases()
    {
        return {
            std::make_tuple(HtmlFixedPageHorizontalAlignment::Center),
            std::make_tuple(HtmlFixedPageHorizontalAlignment::Left),
            std::make_tuple(HtmlFixedPageHorizontalAlignment::Right),
        };
    }
};

TEST_P(HorizontalAlignmentVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->HorizontalAlignment(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlFixedSaveOptions, HorizontalAlignmentVP, ::testing::ValuesIn(HorizontalAlignmentVP::TestCases()));

TEST_F(ExHtmlFixedSaveOptions, PageMargins)
{
    s_instance->PageMargins();
}

TEST_F(ExHtmlFixedSaveOptions, PageMarginsException)
{
    s_instance->PageMarginsException();
}

TEST_F(ExHtmlFixedSaveOptions, OptimizeGraphicsOutput)
{
    s_instance->OptimizeGraphicsOutput();
}

TEST_F(ExHtmlFixedSaveOptions, UsingMachineFonts)
{
    s_instance->UsingMachineFonts();
}

TEST_F(ExHtmlFixedSaveOptions, HtmlFixedResourceFolder)
{
    s_instance->HtmlFixedResourceFolder();
}

}} // namespace ApiExamples::gtest_test
