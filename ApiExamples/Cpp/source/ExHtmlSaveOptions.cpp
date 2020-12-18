// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHtmlSaveOptions.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Saving/ExportListLabels.h>
#include <Aspose.Words.Cpp/Model/Saving/FontSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlOfficeMathOutputMode.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlVersion.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSavingArgs.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples { namespace gtest_test {

class ExHtmlSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExHtmlSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExHtmlSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExHtmlSaveOptions> ExHtmlSaveOptions::s_instance;

using ExportPageMargins_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportPageMargins)>::type;

struct ExportPageMarginsVP : public ExHtmlSaveOptions, public ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExportPageMargins_Args>
{
    static std::vector<ExportPageMargins_Args> TestCases()
    {
        return {
            std::make_tuple(SaveFormat::Html),
            std::make_tuple(SaveFormat::Mhtml),
            std::make_tuple(SaveFormat::Epub),
        };
    }
};

TEST_P(ExportPageMarginsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportPageMargins(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlSaveOptions, ExportPageMarginsVP, ::testing::ValuesIn(ExportPageMarginsVP::TestCases()));

using ExportOfficeMath_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportOfficeMath)>::type;

struct ExportOfficeMathVP : public ExHtmlSaveOptions, public ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExportOfficeMath_Args>
{
    static std::vector<ExportOfficeMath_Args> TestCases()
    {
        return {
            std::make_tuple(SaveFormat::Html, HtmlOfficeMathOutputMode::Image),
            std::make_tuple(SaveFormat::Mhtml, HtmlOfficeMathOutputMode::MathML),
            std::make_tuple(SaveFormat::Epub, HtmlOfficeMathOutputMode::Text),
        };
    }
};

TEST_P(ExportOfficeMathVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportOfficeMath(get<0>(params), get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlSaveOptions, ExportOfficeMathVP, ::testing::ValuesIn(ExportOfficeMathVP::TestCases()));

using ExportTextBoxAsSvg_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportTextBoxAsSvg)>::type;

struct ExportTextBoxAsSvgVP : public ExHtmlSaveOptions, public ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExportTextBoxAsSvg_Args>
{
    static std::vector<ExportTextBoxAsSvg_Args> TestCases()
    {
        return {
            std::make_tuple(SaveFormat::Html, true),
            std::make_tuple(SaveFormat::Epub, true),
            std::make_tuple(SaveFormat::Mhtml, false),
        };
    }
};

TEST_P(ExportTextBoxAsSvgVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportTextBoxAsSvg(get<0>(params), get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlSaveOptions, ExportTextBoxAsSvgVP, ::testing::ValuesIn(ExportTextBoxAsSvgVP::TestCases()));

using ControlListLabelsExport_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ControlListLabelsExport)>::type;

struct ControlListLabelsExportVP : public ExHtmlSaveOptions,
                                   public ApiExamples::ExHtmlSaveOptions,
                                   public ::testing::WithParamInterface<ControlListLabelsExport_Args>
{
    static std::vector<ControlListLabelsExport_Args> TestCases()
    {
        return {
            std::make_tuple(ExportListLabels::Auto),
            std::make_tuple(ExportListLabels::AsInlineText),
            std::make_tuple(ExportListLabels::ByHtmlTags),
        };
    }
};

TEST_P(ControlListLabelsExportVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ControlListLabelsExport(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlSaveOptions, ControlListLabelsExportVP, ::testing::ValuesIn(ControlListLabelsExportVP::TestCases()));

using ExportUrlForLinkedImage_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportUrlForLinkedImage)>::type;

struct ExportUrlForLinkedImageVP : public ExHtmlSaveOptions,
                                   public ApiExamples::ExHtmlSaveOptions,
                                   public ::testing::WithParamInterface<ExportUrlForLinkedImage_Args>
{
    static std::vector<ExportUrlForLinkedImage_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExportUrlForLinkedImageVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportUrlForLinkedImage(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlSaveOptions, ExportUrlForLinkedImageVP, ::testing::ValuesIn(ExportUrlForLinkedImageVP::TestCases()));

TEST_F(ExHtmlSaveOptions, ExportRoundtripInformation)
{
    s_instance->ExportRoundtripInformation();
}

TEST_F(ExHtmlSaveOptions, RoundtripInformationDefaulValue)
{
    s_instance->RoundtripInformationDefaulValue();
}

TEST_F(ExHtmlSaveOptions, ExternalResourceSavingConfig)
{
    s_instance->ExternalResourceSavingConfig();
}

TEST_F(ExHtmlSaveOptions, ConvertFontsAsBase64)
{
    s_instance->ConvertFontsAsBase64();
}

using Html5Support_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::Html5Support)>::type;

struct Html5SupportVP : public ExHtmlSaveOptions, public ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<Html5Support_Args>
{
    static std::vector<Html5Support_Args> TestCases()
    {
        return {
            std::make_tuple(HtmlVersion::Html5),
            std::make_tuple(HtmlVersion::Xhtml),
        };
    }
};

TEST_P(Html5SupportVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Html5Support(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlSaveOptions, Html5SupportVP, ::testing::ValuesIn(Html5SupportVP::TestCases()));

using ExportFonts_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportFonts)>::type;

struct ExportFontsVP : public ExHtmlSaveOptions, public ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExportFonts_Args>
{
    static std::vector<ExportFonts_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExportFontsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportFonts(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlSaveOptions, ExportFontsVP, ::testing::ValuesIn(ExportFontsVP::TestCases()));

TEST_F(ExHtmlSaveOptions, ResourceFolderPriority)
{
    s_instance->ResourceFolderPriority();
}

TEST_F(ExHtmlSaveOptions, ResourceFolderLowPriority)
{
    s_instance->ResourceFolderLowPriority();
}

TEST_F(ExHtmlSaveOptions, SvgMetafileFormat)
{
    s_instance->SvgMetafileFormat();
}

TEST_F(ExHtmlSaveOptions, PngMetafileFormat)
{
    s_instance->PngMetafileFormat();
}

TEST_F(ExHtmlSaveOptions, EmfOrWmfMetafileFormat)
{
    s_instance->EmfOrWmfMetafileFormat();
}

TEST_F(ExHtmlSaveOptions, CssClassNamesPrefix)
{
    s_instance->CssClassNamesPrefix();
}

TEST_F(ExHtmlSaveOptions, CssClassNamesNotValidPrefix)
{
    s_instance->CssClassNamesNotValidPrefix();
}

TEST_F(ExHtmlSaveOptions, CssClassNamesNullPrefix)
{
    s_instance->CssClassNamesNullPrefix();
}

TEST_F(ExHtmlSaveOptions, ContentIdScheme)
{
    s_instance->ContentIdScheme();
}

TEST_F(ExHtmlSaveOptions, HeadingLevels)
{
    s_instance->HeadingLevels();
}

TEST_F(ExHtmlSaveOptions, NegativeIndent)
{
    s_instance->NegativeIndent();
}

TEST_F(ExHtmlSaveOptions, FolderAlias)
{
    s_instance->FolderAlias();
}

TEST_F(ExHtmlSaveOptions, SaveExportedFonts)
{
    s_instance->SaveExportedFonts();
}

TEST_F(ExHtmlSaveOptions, HtmlVersion)
{
    s_instance->HtmlVersion_();
}

TEST_F(ExHtmlSaveOptions, EpubHeadings)
{
    s_instance->EpubHeadings();
}

TEST_F(ExHtmlSaveOptions, Doc2EpubSaveOptions)
{
    s_instance->Doc2EpubSaveOptions();
}

TEST_F(ExHtmlSaveOptions, ContentIdUrls)
{
    s_instance->ContentIdUrls();
}

TEST_F(ExHtmlSaveOptions, DropDownFormField)
{
    s_instance->DropDownFormField();
}

TEST_F(ExHtmlSaveOptions, ExportBase64)
{
    s_instance->ExportBase64();
}

TEST_F(ExHtmlSaveOptions, ExportLanguageInformation)
{
    s_instance->ExportLanguageInformation();
}

TEST_F(ExHtmlSaveOptions, List)
{
    s_instance->List_();
}

TEST_F(ExHtmlSaveOptions, ExportPageMargins2)
{
    s_instance->ExportPageMargins2();
}

TEST_F(ExHtmlSaveOptions, ExportPageSetup)
{
    s_instance->ExportPageSetup();
}

TEST_F(ExHtmlSaveOptions, RelativeFontSize)
{
    s_instance->RelativeFontSize();
}

TEST_F(ExHtmlSaveOptions, ExportTextBox)
{
    s_instance->ExportTextBox();
}

TEST_F(ExHtmlSaveOptions, RoundTripInformation)
{
    s_instance->RoundTripInformation();
}

TEST_F(ExHtmlSaveOptions, ExportTocPageNumbers)
{
    s_instance->ExportTocPageNumbers();
}

TEST_F(ExHtmlSaveOptions, FontSubsetting)
{
    s_instance->FontSubsetting();
}

TEST_F(ExHtmlSaveOptions, MetafileFormat)
{
    s_instance->MetafileFormat();
}

TEST_F(ExHtmlSaveOptions, OfficeMathOutputMode)
{
    s_instance->OfficeMathOutputMode();
}

TEST_F(ExHtmlSaveOptions, ScaleImageToShapeSize)
{
    s_instance->ScaleImageToShapeSize();
}

TEST_F(ExHtmlSaveOptions, ImageFolder)
{
    s_instance->ImageFolder();
}

TEST_F(ExHtmlSaveOptions, ImageSavingCallback)
{
    s_instance->ImageSavingCallback();
}

using PrettyFormat_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::PrettyFormat)>::type;

struct PrettyFormatVP : public ExHtmlSaveOptions, public ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<PrettyFormat_Args>
{
    static std::vector<PrettyFormat_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(PrettyFormatVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PrettyFormat(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlSaveOptions, PrettyFormatVP, ::testing::ValuesIn(PrettyFormatVP::TestCases()));

}} // namespace ApiExamples::gtest_test
