// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExImageSaveOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples { namespace gtest_test {

class ExImageSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExImageSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExImageSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExImageSaveOptions> ExImageSaveOptions::s_instance;

TEST_F(ExImageSaveOptions, OnePage)
{
    s_instance->OnePage();
}

using ExImageSaveOptions_Renderer_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExImageSaveOptions::Renderer)>::type;

struct ExImageSaveOptions_Renderer : public ExImageSaveOptions,
                                     public ApiExamples::ExImageSaveOptions,
                                     public ::testing::WithParamInterface<ExImageSaveOptions_Renderer_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExImageSaveOptions_Renderer, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Renderer(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExImageSaveOptions_Renderer, ::testing::ValuesIn(ExImageSaveOptions_Renderer::TestCases()));

TEST_F(ExImageSaveOptions, PageSet)
{
    s_instance->PageSet_();
}

TEST_F(ExImageSaveOptions, SkipMono_PageByPage)
{
    RecordProperty("category", "SkipMono");

    s_instance->PageByPage();
}

using ExImageSaveOptions_ColorMode_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExImageSaveOptions::ColorMode)>::type;

struct ExImageSaveOptions_ColorMode : public ExImageSaveOptions,
                                      public ApiExamples::ExImageSaveOptions,
                                      public ::testing::WithParamInterface<ExImageSaveOptions_ColorMode_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(ImageColorMode::Grayscale),
            std::make_tuple(ImageColorMode::None),
        };
    }
};

TEST_P(ExImageSaveOptions_ColorMode, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ColorMode(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExImageSaveOptions_ColorMode, ::testing::ValuesIn(ExImageSaveOptions_ColorMode::TestCases()));

TEST_F(ExImageSaveOptions, PaperColor)
{
    s_instance->PaperColor();
}

using ExImageSaveOptions_PixelFormat_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExImageSaveOptions::PixelFormat)>::type;

struct ExImageSaveOptions_PixelFormat : public ExImageSaveOptions,
                                        public ApiExamples::ExImageSaveOptions,
                                        public ::testing::WithParamInterface<ExImageSaveOptions_PixelFormat_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(ImagePixelFormat::Format1bppIndexed), 
            std::make_tuple(ImagePixelFormat::Format24BppRgb),    
            std::make_tuple(ImagePixelFormat::Format32BppRgb),
        };
    }
};

TEST_P(ExImageSaveOptions_PixelFormat, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PixelFormat(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExImageSaveOptions_PixelFormat, ::testing::ValuesIn(ExImageSaveOptions_PixelFormat::TestCases()));

TEST_F(ExImageSaveOptions, SkipMono_FloydSteinbergDithering)
{
    RecordProperty("category", "SkipMono");

    s_instance->FloydSteinbergDithering();
}

TEST_F(ExImageSaveOptions, EditImage)
{
    s_instance->EditImage();
}

TEST_F(ExImageSaveOptions, JpegQuality)
{
    s_instance->JpegQuality();
}

TEST_F(ExImageSaveOptions, SkipMono_SaveToTiffDefault)
{
    RecordProperty("category", "SkipMono");

    s_instance->SaveToTiffDefault();
}

using ExImageSaveOptions_SkipMono_SkipMono_SkipMono_SkipMono_SkipMono_TiffImageCompression_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExImageSaveOptions::TiffImageCompression)>::type;

struct ExImageSaveOptions_SkipMono_SkipMono_SkipMono_SkipMono_SkipMono_TiffImageCompression
    : public ExImageSaveOptions,
      public ApiExamples::ExImageSaveOptions,
      public ::testing::WithParamInterface<ExImageSaveOptions_SkipMono_SkipMono_SkipMono_SkipMono_SkipMono_TiffImageCompression_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(TiffCompression::None),   std::make_tuple(TiffCompression::Rle),    std::make_tuple(TiffCompression::Lzw),
            std::make_tuple(TiffCompression::Ccitt3), std::make_tuple(TiffCompression::Ccitt4),
        };
    }
};

TEST_P(ExImageSaveOptions_SkipMono_SkipMono_SkipMono_SkipMono_SkipMono_TiffImageCompression, Test)
{
    RecordProperty("category", "SkipMono");
    RecordProperty("category1", "SkipMono");
    RecordProperty("category2", "SkipMono");
    RecordProperty("category3", "SkipMono");
    RecordProperty("category4", "SkipMono");
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TiffImageCompression(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExImageSaveOptions_SkipMono_SkipMono_SkipMono_SkipMono_SkipMono_TiffImageCompression,
                         ::testing::ValuesIn(ExImageSaveOptions_SkipMono_SkipMono_SkipMono_SkipMono_SkipMono_TiffImageCompression::TestCases()));

TEST_F(ExImageSaveOptions, Resolution)
{
    s_instance->Resolution();
}

TEST_F(ExImageSaveOptions, ExportVariousPageRanges)
{
    s_instance->ExportVariousPageRanges();
}

TEST_F(ExImageSaveOptions, RenderInkObject)
{
    s_instance->RenderInkObject();
}

}} // namespace ApiExamples::gtest_test
