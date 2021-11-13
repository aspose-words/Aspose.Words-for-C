// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: Aspose.Pdf is not supported
#include "ExPdfSaveOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
using namespace Aspose::Words::DigitalSignatures;
namespace ApiExamples { namespace gtest_test {

class ExPdfSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExPdfSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExPdfSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExPdfSaveOptions> ExPdfSaveOptions::s_instance;

TEST_F(ExPdfSaveOptions, OnePage)
{
    s_instance->OnePage();
}

TEST_F(ExPdfSaveOptions, HeadingsOutlineLevels)
{
    s_instance->HeadingsOutlineLevels();
}

using ExPdfSaveOptions_CreateMissingOutlineLevels_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::CreateMissingOutlineLevels)>::type;

struct ExPdfSaveOptions_CreateMissingOutlineLevels : public ExPdfSaveOptions,
                                                     public ApiExamples::ExPdfSaveOptions,
                                                     public ::testing::WithParamInterface<ExPdfSaveOptions_CreateMissingOutlineLevels_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_CreateMissingOutlineLevels, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->CreateMissingOutlineLevels(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_CreateMissingOutlineLevels, ::testing::ValuesIn(ExPdfSaveOptions_CreateMissingOutlineLevels::TestCases()));

using ExPdfSaveOptions_TableHeadingOutlines_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::TableHeadingOutlines)>::type;

struct ExPdfSaveOptions_TableHeadingOutlines : public ExPdfSaveOptions,
                                               public ApiExamples::ExPdfSaveOptions,
                                               public ::testing::WithParamInterface<ExPdfSaveOptions_TableHeadingOutlines_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_TableHeadingOutlines, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TableHeadingOutlines(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_TableHeadingOutlines, ::testing::ValuesIn(ExPdfSaveOptions_TableHeadingOutlines::TestCases()));

TEST_F(ExPdfSaveOptions, ExpandedOutlineLevels)
{
    s_instance->ExpandedOutlineLevels();
}

using ExPdfSaveOptions_UpdateFields_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::UpdateFields)>::type;

struct ExPdfSaveOptions_UpdateFields : public ExPdfSaveOptions,
                                       public ApiExamples::ExPdfSaveOptions,
                                       public ::testing::WithParamInterface<ExPdfSaveOptions_UpdateFields_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_UpdateFields, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UpdateFields(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_UpdateFields, ::testing::ValuesIn(ExPdfSaveOptions_UpdateFields::TestCases()));

using ExPdfSaveOptions_PreserveFormFields_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::PreserveFormFields)>::type;

struct ExPdfSaveOptions_PreserveFormFields : public ExPdfSaveOptions,
                                             public ApiExamples::ExPdfSaveOptions,
                                             public ::testing::WithParamInterface<ExPdfSaveOptions_PreserveFormFields_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_PreserveFormFields, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PreserveFormFields(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_PreserveFormFields, ::testing::ValuesIn(ExPdfSaveOptions_PreserveFormFields::TestCases()));

using ExPdfSaveOptions_Compliance_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::Compliance)>::type;

struct ExPdfSaveOptions_Compliance : public ExPdfSaveOptions,
                                     public ApiExamples::ExPdfSaveOptions,
                                     public ::testing::WithParamInterface<ExPdfSaveOptions_Compliance_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(PdfCompliance::PdfA2u),
            std::make_tuple(PdfCompliance::Pdf17),
            std::make_tuple(PdfCompliance::PdfA2a),
        };
    }
};

TEST_P(ExPdfSaveOptions_Compliance, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Compliance(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_Compliance, ::testing::ValuesIn(ExPdfSaveOptions_Compliance::TestCases()));

using ExPdfSaveOptions_TextCompression_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::TextCompression)>::type;

struct ExPdfSaveOptions_TextCompression : public ExPdfSaveOptions,
                                          public ApiExamples::ExPdfSaveOptions,
                                          public ::testing::WithParamInterface<ExPdfSaveOptions_TextCompression_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(PdfTextCompression::None),
            std::make_tuple(PdfTextCompression::Flate),
        };
    }
};

TEST_P(ExPdfSaveOptions_TextCompression, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TextCompression(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_TextCompression, ::testing::ValuesIn(ExPdfSaveOptions_TextCompression::TestCases()));

using ExPdfSaveOptions_ImageCompression_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::ImageCompression)>::type;

struct ExPdfSaveOptions_ImageCompression : public ExPdfSaveOptions,
                                           public ApiExamples::ExPdfSaveOptions,
                                           public ::testing::WithParamInterface<ExPdfSaveOptions_ImageCompression_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(PdfImageCompression::Auto),
            std::make_tuple(PdfImageCompression::Jpeg),
        };
    }
};

TEST_P(ExPdfSaveOptions_ImageCompression, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ImageCompression(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_ImageCompression, ::testing::ValuesIn(ExPdfSaveOptions_ImageCompression::TestCases()));

using ExPdfSaveOptions_ImageColorSpaceExportMode_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::ImageColorSpaceExportMode)>::type;

struct ExPdfSaveOptions_ImageColorSpaceExportMode : public ExPdfSaveOptions,
                                                    public ApiExamples::ExPdfSaveOptions,
                                                    public ::testing::WithParamInterface<ExPdfSaveOptions_ImageColorSpaceExportMode_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(PdfImageColorSpaceExportMode::Auto),
            std::make_tuple(PdfImageColorSpaceExportMode::SimpleCmyk),
        };
    }
};

TEST_P(ExPdfSaveOptions_ImageColorSpaceExportMode, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ImageColorSpaceExportMode(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_ImageColorSpaceExportMode, ::testing::ValuesIn(ExPdfSaveOptions_ImageColorSpaceExportMode::TestCases()));

TEST_F(ExPdfSaveOptions, DownsampleOptions)
{
    s_instance->DownsampleOptions_();
}

using ExPdfSaveOptions_ColorRendering_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::ColorRendering)>::type;

struct ExPdfSaveOptions_ColorRendering : public ExPdfSaveOptions,
                                         public ApiExamples::ExPdfSaveOptions,
                                         public ::testing::WithParamInterface<ExPdfSaveOptions_ColorRendering_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(ColorMode::Grayscale),
            std::make_tuple(ColorMode::Normal),
        };
    }
};

TEST_P(ExPdfSaveOptions_ColorRendering, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ColorRendering(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_ColorRendering, ::testing::ValuesIn(ExPdfSaveOptions_ColorRendering::TestCases()));

using ExPdfSaveOptions_DocTitle_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::DocTitle)>::type;

struct ExPdfSaveOptions_DocTitle : public ExPdfSaveOptions,
                                   public ApiExamples::ExPdfSaveOptions,
                                   public ::testing::WithParamInterface<ExPdfSaveOptions_DocTitle_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_DocTitle, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DocTitle(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_DocTitle, ::testing::ValuesIn(ExPdfSaveOptions_DocTitle::TestCases()));

using ExPdfSaveOptions_MemoryOptimization_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::MemoryOptimization)>::type;

struct ExPdfSaveOptions_MemoryOptimization : public ExPdfSaveOptions,
                                             public ApiExamples::ExPdfSaveOptions,
                                             public ::testing::WithParamInterface<ExPdfSaveOptions_MemoryOptimization_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_MemoryOptimization, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MemoryOptimization(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_MemoryOptimization, ::testing::ValuesIn(ExPdfSaveOptions_MemoryOptimization::TestCases()));

using ExPdfSaveOptions_EscapeUri_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::EscapeUri)>::type;

struct ExPdfSaveOptions_EscapeUri : public ExPdfSaveOptions,
                                    public ApiExamples::ExPdfSaveOptions,
                                    public ::testing::WithParamInterface<ExPdfSaveOptions_EscapeUri_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(u"https://www.google.com/search?q= aspose", u"https://www.google.com/search?q=%20aspose"),
            std::make_tuple(u"https://www.google.com/search?q=%20aspose", u"https://www.google.com/search?q=%20aspose"),
        };
    }
};

TEST_P(ExPdfSaveOptions_EscapeUri, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->EscapeUri(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_EscapeUri, ::testing::ValuesIn(ExPdfSaveOptions_EscapeUri::TestCases()));

using ExPdfSaveOptions_OpenHyperlinksInNewWindow_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::OpenHyperlinksInNewWindow)>::type;

struct ExPdfSaveOptions_OpenHyperlinksInNewWindow : public ExPdfSaveOptions,
                                                    public ApiExamples::ExPdfSaveOptions,
                                                    public ::testing::WithParamInterface<ExPdfSaveOptions_OpenHyperlinksInNewWindow_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_OpenHyperlinksInNewWindow, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->OpenHyperlinksInNewWindow(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_OpenHyperlinksInNewWindow, ::testing::ValuesIn(ExPdfSaveOptions_OpenHyperlinksInNewWindow::TestCases()));

TEST_F(ExPdfSaveOptions, SkipMono_HandleBinaryRasterWarnings)
{
    RecordProperty("category", "SkipMono");

    s_instance->HandleBinaryRasterWarnings();
}

using ExPdfSaveOptions_HeaderFooterBookmarksExportMode_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::HeaderFooterBookmarksExportMode)>::type;

struct ExPdfSaveOptions_HeaderFooterBookmarksExportMode : public ExPdfSaveOptions,
                                                          public ApiExamples::ExPdfSaveOptions,
                                                          public ::testing::WithParamInterface<ExPdfSaveOptions_HeaderFooterBookmarksExportMode_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(Aspose::Words::Saving::HeaderFooterBookmarksExportMode::None),
            std::make_tuple(Aspose::Words::Saving::HeaderFooterBookmarksExportMode::First),
            std::make_tuple(Aspose::Words::Saving::HeaderFooterBookmarksExportMode::All),
        };
    }
};

TEST_P(ExPdfSaveOptions_HeaderFooterBookmarksExportMode, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->HeaderFooterBookmarksExportMode(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_HeaderFooterBookmarksExportMode,
                         ::testing::ValuesIn(ExPdfSaveOptions_HeaderFooterBookmarksExportMode::TestCases()));

TEST_F(ExPdfSaveOptions, UnsupportedImageFormatWarning)
{
    s_instance->UnsupportedImageFormatWarning();
}

using ExPdfSaveOptions_FontsScaledToMetafileSize_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::FontsScaledToMetafileSize)>::type;

struct ExPdfSaveOptions_FontsScaledToMetafileSize : public ExPdfSaveOptions,
                                                    public ApiExamples::ExPdfSaveOptions,
                                                    public ::testing::WithParamInterface<ExPdfSaveOptions_FontsScaledToMetafileSize_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_FontsScaledToMetafileSize, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FontsScaledToMetafileSize(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_FontsScaledToMetafileSize, ::testing::ValuesIn(ExPdfSaveOptions_FontsScaledToMetafileSize::TestCases()));

using ExPdfSaveOptions_EmbedFullFonts_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::EmbedFullFonts)>::type;

struct ExPdfSaveOptions_EmbedFullFonts : public ExPdfSaveOptions,
                                         public ApiExamples::ExPdfSaveOptions,
                                         public ::testing::WithParamInterface<ExPdfSaveOptions_EmbedFullFonts_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_EmbedFullFonts, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->EmbedFullFonts(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_EmbedFullFonts, ::testing::ValuesIn(ExPdfSaveOptions_EmbedFullFonts::TestCases()));

using ExPdfSaveOptions_EmbedWindowsFonts_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::EmbedWindowsFonts)>::type;

struct ExPdfSaveOptions_EmbedWindowsFonts : public ExPdfSaveOptions,
                                            public ApiExamples::ExPdfSaveOptions,
                                            public ::testing::WithParamInterface<ExPdfSaveOptions_EmbedWindowsFonts_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(PdfFontEmbeddingMode::EmbedAll),
            std::make_tuple(PdfFontEmbeddingMode::EmbedNone),
            std::make_tuple(PdfFontEmbeddingMode::EmbedNonstandard),
        };
    }
};

TEST_P(ExPdfSaveOptions_EmbedWindowsFonts, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->EmbedWindowsFonts(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_EmbedWindowsFonts, ::testing::ValuesIn(ExPdfSaveOptions_EmbedWindowsFonts::TestCases()));

using ExPdfSaveOptions_EmbedCoreFonts_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::EmbedCoreFonts)>::type;

struct ExPdfSaveOptions_EmbedCoreFonts : public ExPdfSaveOptions,
                                         public ApiExamples::ExPdfSaveOptions,
                                         public ::testing::WithParamInterface<ExPdfSaveOptions_EmbedCoreFonts_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_EmbedCoreFonts, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->EmbedCoreFonts(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_EmbedCoreFonts, ::testing::ValuesIn(ExPdfSaveOptions_EmbedCoreFonts::TestCases()));

using ExPdfSaveOptions_AdditionalTextPositioning_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::AdditionalTextPositioning)>::type;

struct ExPdfSaveOptions_AdditionalTextPositioning : public ExPdfSaveOptions,
                                                    public ApiExamples::ExPdfSaveOptions,
                                                    public ::testing::WithParamInterface<ExPdfSaveOptions_AdditionalTextPositioning_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_AdditionalTextPositioning, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AdditionalTextPositioning(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_AdditionalTextPositioning, ::testing::ValuesIn(ExPdfSaveOptions_AdditionalTextPositioning::TestCases()));

using ExPdfSaveOptions_SaveAsPdfBookFold_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::SaveAsPdfBookFold)>::type;

struct ExPdfSaveOptions_SaveAsPdfBookFold : public ExPdfSaveOptions,
                                            public ApiExamples::ExPdfSaveOptions,
                                            public ::testing::WithParamInterface<ExPdfSaveOptions_SaveAsPdfBookFold_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_SaveAsPdfBookFold, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SaveAsPdfBookFold(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_SaveAsPdfBookFold, ::testing::ValuesIn(ExPdfSaveOptions_SaveAsPdfBookFold::TestCases()));

TEST_F(ExPdfSaveOptions, ZoomBehaviour)
{
    s_instance->ZoomBehaviour();
}

using ExPdfSaveOptions_PageMode_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::PageMode)>::type;

struct ExPdfSaveOptions_PageMode : public ExPdfSaveOptions,
                                   public ApiExamples::ExPdfSaveOptions,
                                   public ::testing::WithParamInterface<ExPdfSaveOptions_PageMode_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(PdfPageMode::FullScreen),  std::make_tuple(PdfPageMode::UseThumbs), std::make_tuple(PdfPageMode::UseOC),
            std::make_tuple(PdfPageMode::UseOutlines), std::make_tuple(PdfPageMode::UseNone),
        };
    }
};

TEST_P(ExPdfSaveOptions_PageMode, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PageMode(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_PageMode, ::testing::ValuesIn(ExPdfSaveOptions_PageMode::TestCases()));

using ExPdfSaveOptions_NoteHyperlinks_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::NoteHyperlinks)>::type;

struct ExPdfSaveOptions_NoteHyperlinks : public ExPdfSaveOptions,
                                         public ApiExamples::ExPdfSaveOptions,
                                         public ::testing::WithParamInterface<ExPdfSaveOptions_NoteHyperlinks_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_NoteHyperlinks, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->NoteHyperlinks(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_NoteHyperlinks, ::testing::ValuesIn(ExPdfSaveOptions_NoteHyperlinks::TestCases()));

using ExPdfSaveOptions_CustomPropertiesExport_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::CustomPropertiesExport)>::type;

struct ExPdfSaveOptions_CustomPropertiesExport : public ExPdfSaveOptions,
                                                 public ApiExamples::ExPdfSaveOptions,
                                                 public ::testing::WithParamInterface<ExPdfSaveOptions_CustomPropertiesExport_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(PdfCustomPropertiesExport::None),
            std::make_tuple(PdfCustomPropertiesExport::Standard),
            std::make_tuple(PdfCustomPropertiesExport::Metadata),
        };
    }
};

TEST_P(ExPdfSaveOptions_CustomPropertiesExport, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->CustomPropertiesExport(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_CustomPropertiesExport, ::testing::ValuesIn(ExPdfSaveOptions_CustomPropertiesExport::TestCases()));

using ExPdfSaveOptions_DrawingMLEffects_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::DrawingMLEffects)>::type;

struct ExPdfSaveOptions_DrawingMLEffects : public ExPdfSaveOptions,
                                           public ApiExamples::ExPdfSaveOptions,
                                           public ::testing::WithParamInterface<ExPdfSaveOptions_DrawingMLEffects_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(DmlEffectsRenderingMode::None),
            std::make_tuple(DmlEffectsRenderingMode::Simplified),
            std::make_tuple(DmlEffectsRenderingMode::Fine),
        };
    }
};

TEST_P(ExPdfSaveOptions_DrawingMLEffects, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DrawingMLEffects(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_DrawingMLEffects, ::testing::ValuesIn(ExPdfSaveOptions_DrawingMLEffects::TestCases()));

using ExPdfSaveOptions_DrawingMLFallback_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::DrawingMLFallback)>::type;

struct ExPdfSaveOptions_DrawingMLFallback : public ExPdfSaveOptions,
                                            public ApiExamples::ExPdfSaveOptions,
                                            public ::testing::WithParamInterface<ExPdfSaveOptions_DrawingMLFallback_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(DmlRenderingMode::Fallback),
            std::make_tuple(DmlRenderingMode::DrawingML),
        };
    }
};

TEST_P(ExPdfSaveOptions_DrawingMLFallback, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DrawingMLFallback(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_DrawingMLFallback, ::testing::ValuesIn(ExPdfSaveOptions_DrawingMLFallback::TestCases()));

using ExPdfSaveOptions_ExportDocumentStructure_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::ExportDocumentStructure)>::type;

struct ExPdfSaveOptions_ExportDocumentStructure : public ExPdfSaveOptions,
                                                  public ApiExamples::ExPdfSaveOptions,
                                                  public ::testing::WithParamInterface<ExPdfSaveOptions_ExportDocumentStructure_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPdfSaveOptions_ExportDocumentStructure, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportDocumentStructure(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_ExportDocumentStructure, ::testing::ValuesIn(ExPdfSaveOptions_ExportDocumentStructure::TestCases()));

TEST_F(ExPdfSaveOptions, PdfDigitalSignature)
{
    s_instance->PdfDigitalSignature();
}

TEST_F(ExPdfSaveOptions, PdfDigitalSignatureTimestamp)
{
    s_instance->PdfDigitalSignatureTimestamp();
}

using ExPdfSaveOptions_RenderMetafile_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::RenderMetafile)>::type;

struct ExPdfSaveOptions_RenderMetafile : public ExPdfSaveOptions,
                                         public ApiExamples::ExPdfSaveOptions,
                                         public ::testing::WithParamInterface<ExPdfSaveOptions_RenderMetafile_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(EmfPlusDualRenderingMode::Emf),
            std::make_tuple(EmfPlusDualRenderingMode::EmfPlus),
            std::make_tuple(EmfPlusDualRenderingMode::EmfPlusWithFallback),
        };
    }
};

TEST_P(ExPdfSaveOptions_RenderMetafile, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RenderMetafile(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_RenderMetafile, ::testing::ValuesIn(ExPdfSaveOptions_RenderMetafile::TestCases()));

TEST_F(ExPdfSaveOptions, EncryptionPermissions)
{
    s_instance->EncryptionPermissions();
}

using ExPdfSaveOptions_SetNumeralFormat_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::SetNumeralFormat)>::type;

struct ExPdfSaveOptions_SetNumeralFormat : public ExPdfSaveOptions,
                                           public ApiExamples::ExPdfSaveOptions,
                                           public ::testing::WithParamInterface<ExPdfSaveOptions_SetNumeralFormat_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(NumeralFormat::ArabicIndic), std::make_tuple(NumeralFormat::Context), std::make_tuple(NumeralFormat::EasternArabicIndic),
            std::make_tuple(NumeralFormat::European),    std::make_tuple(NumeralFormat::System),
        };
    }
};

TEST_P(ExPdfSaveOptions_SetNumeralFormat, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SetNumeralFormat(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPdfSaveOptions_SetNumeralFormat, ::testing::ValuesIn(ExPdfSaveOptions_SetNumeralFormat::TestCases()));

TEST_F(ExPdfSaveOptions, ExportPageSet)
{
    s_instance->ExportPageSet();
}

TEST_F(ExPdfSaveOptions, ExportLanguageToSpanTag)
{
    s_instance->ExportLanguageToSpanTag();
}

}} // namespace ApiExamples::gtest_test
