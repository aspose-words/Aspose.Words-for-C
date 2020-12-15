// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: Aspose.Pdf is not supported
#include "ExPdfSaveOptions.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Document/WarningSource.h>
#include <Aspose.Words.Cpp/Model/Document/WarningType.h>
#include <Aspose.Words.Cpp/Model/Saving/DmlEffectsRenderingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/DmlRenderingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/EmfPlusDualRenderingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/HeaderFooterBookmarksExportMode.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfCustomPropertiesExport.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfPageMode.h>
#include <system/collections/list.h>
#include <system/enum_helpers.h>
#include <system/shared_ptr.h>
#include <system/string.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
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

TEST_F(ExPdfSaveOptions, CreateMissingOutlineLevels)
{
    s_instance->CreateMissingOutlineLevels();
}

TEST_F(ExPdfSaveOptions, TableHeadingOutlines)
{
    s_instance->TableHeadingOutlines();
}

using UpdateFields_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::UpdateFields)>::type;

struct UpdateFieldsVP : public ExPdfSaveOptions, public ApiExamples::ExPdfSaveOptions, public ::testing::WithParamInterface<UpdateFields_Args>
{
    static std::vector<UpdateFields_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(UpdateFieldsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UpdateFields(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, UpdateFieldsVP, ::testing::ValuesIn(UpdateFieldsVP::TestCases()));

TEST_F(ExPdfSaveOptions, ImageCompression)
{
    s_instance->ImageCompression();
}

TEST_F(ExPdfSaveOptions, DownsampleOptions)
{
    s_instance->DownsampleOptions_();
}

TEST_F(ExPdfSaveOptions, ColorRendering)
{
    s_instance->ColorRendering();
}

TEST_F(ExPdfSaveOptions, WindowsBarPdfTitle)
{
    s_instance->WindowsBarPdfTitle();
}

TEST_F(ExPdfSaveOptions, MemoryOptimization)
{
    s_instance->MemoryOptimization();
}

using EscapeUri_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::EscapeUri)>::type;

struct EscapeUriVP : public ExPdfSaveOptions, public ApiExamples::ExPdfSaveOptions, public ::testing::WithParamInterface<EscapeUri_Args>
{
    static std::vector<EscapeUri_Args> TestCases()
    {
        return {
            std::make_tuple(u"https://www.google.com/search?q= aspose", u"app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);", true),
            std::make_tuple(u"https://www.google.com/search?q=%20aspose", u"app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);", true),
            std::make_tuple(u"https://www.google.com/search?q= aspose", u"app.launchURL(\"https://www.google.com/search?q= aspose\", true);", false),
            std::make_tuple(u"https://www.google.com/search?q=%20aspose", u"app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);", false),
        };
    }
};

TEST_P(EscapeUriVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->EscapeUri(get<0>(params), get<1>(params), get<2>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, EscapeUriVP, ::testing::ValuesIn(EscapeUriVP::TestCases()));

TEST_F(ExPdfSaveOptions, SkipMono_HandleBinaryRasterWarnings)
{
    RecordProperty("category", "SkipMono");

    s_instance->HandleBinaryRasterWarnings();
}

using HeaderFooterBookmarksExportMode_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::HeaderFooterBookmarksExportMode)>::type;

struct HeaderFooterBookmarksExportModeVP : public ExPdfSaveOptions,
                                           public ApiExamples::ExPdfSaveOptions,
                                           public ::testing::WithParamInterface<HeaderFooterBookmarksExportMode_Args>
{
    static std::vector<HeaderFooterBookmarksExportMode_Args> TestCases()
    {
        return {
            std::make_tuple(Aspose::Words::Saving::HeaderFooterBookmarksExportMode::None),
            std::make_tuple(Aspose::Words::Saving::HeaderFooterBookmarksExportMode::First),
            std::make_tuple(Aspose::Words::Saving::HeaderFooterBookmarksExportMode::All),
        };
    }
};

TEST_P(HeaderFooterBookmarksExportModeVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->HeaderFooterBookmarksExportMode(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, HeaderFooterBookmarksExportModeVP, ::testing::ValuesIn(HeaderFooterBookmarksExportModeVP::TestCases()));

TEST_F(ExPdfSaveOptions, UnsupportedImageFormatWarning)
{
    s_instance->UnsupportedImageFormatWarning();
}

using FontsScaledToMetafileSize_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::FontsScaledToMetafileSize)>::type;

struct FontsScaledToMetafileSizeVP : public ExPdfSaveOptions,
                                     public ApiExamples::ExPdfSaveOptions,
                                     public ::testing::WithParamInterface<FontsScaledToMetafileSize_Args>
{
    static std::vector<FontsScaledToMetafileSize_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(FontsScaledToMetafileSizeVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FontsScaledToMetafileSize(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, FontsScaledToMetafileSizeVP, ::testing::ValuesIn(FontsScaledToMetafileSizeVP::TestCases()));

using AdditionalTextPositioning_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::AdditionalTextPositioning)>::type;

struct AdditionalTextPositioningVP : public ExPdfSaveOptions,
                                     public ApiExamples::ExPdfSaveOptions,
                                     public ::testing::WithParamInterface<AdditionalTextPositioning_Args>
{
    static std::vector<AdditionalTextPositioning_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(AdditionalTextPositioningVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AdditionalTextPositioning(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, AdditionalTextPositioningVP, ::testing::ValuesIn(AdditionalTextPositioningVP::TestCases()));

using SaveAsPdfBookFold_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::SaveAsPdfBookFold)>::type;

struct SaveAsPdfBookFoldVP : public ExPdfSaveOptions, public ApiExamples::ExPdfSaveOptions, public ::testing::WithParamInterface<SaveAsPdfBookFold_Args>
{
    static std::vector<SaveAsPdfBookFold_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(SaveAsPdfBookFoldVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SaveAsPdfBookFold(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, SaveAsPdfBookFoldVP, ::testing::ValuesIn(SaveAsPdfBookFoldVP::TestCases()));

TEST_F(ExPdfSaveOptions, ZoomBehaviour)
{
    s_instance->ZoomBehaviour();
}

using PageMode_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::PageMode)>::type;

struct PageModeVP : public ExPdfSaveOptions, public ApiExamples::ExPdfSaveOptions, public ::testing::WithParamInterface<PageMode_Args>
{
    static std::vector<PageMode_Args> TestCases()
    {
        return {
            std::make_tuple(PdfPageMode::FullScreen),  std::make_tuple(PdfPageMode::UseThumbs), std::make_tuple(PdfPageMode::UseOC),
            std::make_tuple(PdfPageMode::UseOutlines), std::make_tuple(PdfPageMode::UseNone),
        };
    }
};

TEST_P(PageModeVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PageMode(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, PageModeVP, ::testing::ValuesIn(PageModeVP::TestCases()));

using NoteHyperlinks_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::NoteHyperlinks)>::type;

struct NoteHyperlinksVP : public ExPdfSaveOptions, public ApiExamples::ExPdfSaveOptions, public ::testing::WithParamInterface<NoteHyperlinks_Args>
{
    static std::vector<NoteHyperlinks_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(NoteHyperlinksVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->NoteHyperlinks(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, NoteHyperlinksVP, ::testing::ValuesIn(NoteHyperlinksVP::TestCases()));

using CustomPropertiesExport_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::CustomPropertiesExport)>::type;

struct CustomPropertiesExportVP : public ExPdfSaveOptions,
                                  public ApiExamples::ExPdfSaveOptions,
                                  public ::testing::WithParamInterface<CustomPropertiesExport_Args>
{
    static std::vector<CustomPropertiesExport_Args> TestCases()
    {
        return {
            std::make_tuple(PdfCustomPropertiesExport::None),
            std::make_tuple(PdfCustomPropertiesExport::Standard),
            std::make_tuple(PdfCustomPropertiesExport::Metadata),
        };
    }
};

TEST_P(CustomPropertiesExportVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->CustomPropertiesExport(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, CustomPropertiesExportVP, ::testing::ValuesIn(CustomPropertiesExportVP::TestCases()));

using DrawingMLEffects_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::DrawingMLEffects)>::type;

struct DrawingMLEffectsVP : public ExPdfSaveOptions, public ApiExamples::ExPdfSaveOptions, public ::testing::WithParamInterface<DrawingMLEffects_Args>
{
    static std::vector<DrawingMLEffects_Args> TestCases()
    {
        return {
            std::make_tuple(DmlEffectsRenderingMode::None),
            std::make_tuple(DmlEffectsRenderingMode::Simplified),
            std::make_tuple(DmlEffectsRenderingMode::Fine),
        };
    }
};

TEST_P(DrawingMLEffectsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DrawingMLEffects(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, DrawingMLEffectsVP, ::testing::ValuesIn(DrawingMLEffectsVP::TestCases()));

using DrawingMLFallback_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::DrawingMLFallback)>::type;

struct DrawingMLFallbackVP : public ExPdfSaveOptions, public ApiExamples::ExPdfSaveOptions, public ::testing::WithParamInterface<DrawingMLFallback_Args>
{
    static std::vector<DrawingMLFallback_Args> TestCases()
    {
        return {
            std::make_tuple(DmlRenderingMode::Fallback),
            std::make_tuple(DmlRenderingMode::DrawingML),
        };
    }
};

TEST_P(DrawingMLFallbackVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DrawingMLFallback(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, DrawingMLFallbackVP, ::testing::ValuesIn(DrawingMLFallbackVP::TestCases()));

using ExportDocumentStructure_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::ExportDocumentStructure)>::type;

struct ExportDocumentStructureVP : public ExPdfSaveOptions,
                                   public ApiExamples::ExPdfSaveOptions,
                                   public ::testing::WithParamInterface<ExportDocumentStructure_Args>
{
    static std::vector<ExportDocumentStructure_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExportDocumentStructureVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportDocumentStructure(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, ExportDocumentStructureVP, ::testing::ValuesIn(ExportDocumentStructureVP::TestCases()));

using PreblendImages_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::PreblendImages)>::type;

struct PreblendImagesVP : public ExPdfSaveOptions, public ApiExamples::ExPdfSaveOptions, public ::testing::WithParamInterface<PreblendImages_Args>
{
    static std::vector<PreblendImages_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(PreblendImagesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PreblendImages(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, PreblendImagesVP, ::testing::ValuesIn(PreblendImagesVP::TestCases()));

TEST_F(ExPdfSaveOptions, InterpolateImages)
{
    s_instance->InterpolateImages();
}

TEST_F(ExPdfSaveOptions, SkipMono_Dml3DEffectsRenderingModeTest)
{
    RecordProperty("category", "SkipMono");

    s_instance->Dml3DEffectsRenderingModeTest();
}

TEST_F(ExPdfSaveOptions, PdfDigitalSignature)
{
    s_instance->PdfDigitalSignature();
}

TEST_F(ExPdfSaveOptions, PdfDigitalSignatureTimestamp)
{
    s_instance->PdfDigitalSignatureTimestamp();
}

using RenderMetafile_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPdfSaveOptions::RenderMetafile)>::type;

struct RenderMetafileVP : public ExPdfSaveOptions, public ApiExamples::ExPdfSaveOptions, public ::testing::WithParamInterface<RenderMetafile_Args>
{
    static std::vector<RenderMetafile_Args> TestCases()
    {
        return {
            std::make_tuple(EmfPlusDualRenderingMode::Emf),
            std::make_tuple(EmfPlusDualRenderingMode::EmfPlus),
            std::make_tuple(EmfPlusDualRenderingMode::EmfPlusWithFallback),
        };
    }
};

TEST_P(RenderMetafileVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RenderMetafile(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExPdfSaveOptions, RenderMetafileVP, ::testing::ValuesIn(RenderMetafileVP::TestCases()));

}} // namespace ApiExamples::gtest_test
