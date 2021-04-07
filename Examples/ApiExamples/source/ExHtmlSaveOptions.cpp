// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHtmlSaveOptions.h"

#include <system/primitive_types.h>
#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
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

using ExHtmlSaveOptions_ExportPageMarginsEpub_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportPageMarginsEpub)>::type;

struct ExHtmlSaveOptions_ExportPageMarginsEpub : public ExHtmlSaveOptions,
                                                 public ApiExamples::ExHtmlSaveOptions,
                                                 public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportPageMarginsEpub_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(SaveFormat::Html),
            std::make_tuple(SaveFormat::Mhtml),
            std::make_tuple(SaveFormat::Epub),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportPageMarginsEpub, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportPageMarginsEpub(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportPageMarginsEpub, ::testing::ValuesIn(ExHtmlSaveOptions_ExportPageMarginsEpub::TestCases()));

using ExHtmlSaveOptions_ExportOfficeMathEpub_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportOfficeMathEpub)>::type;

struct ExHtmlSaveOptions_ExportOfficeMathEpub : public ExHtmlSaveOptions,
                                                public ApiExamples::ExHtmlSaveOptions,
                                                public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportOfficeMathEpub_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(SaveFormat::Html, HtmlOfficeMathOutputMode::Image),
            std::make_tuple(SaveFormat::Mhtml, HtmlOfficeMathOutputMode::MathML),
            std::make_tuple(SaveFormat::Epub, HtmlOfficeMathOutputMode::Text),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportOfficeMathEpub, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportOfficeMathEpub(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportOfficeMathEpub, ::testing::ValuesIn(ExHtmlSaveOptions_ExportOfficeMathEpub::TestCases()));

using ExHtmlSaveOptions_ExportTextBoxAsSvgEpub_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportTextBoxAsSvgEpub)>::type;

struct ExHtmlSaveOptions_ExportTextBoxAsSvgEpub : public ExHtmlSaveOptions,
                                                  public ApiExamples::ExHtmlSaveOptions,
                                                  public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportTextBoxAsSvgEpub_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(SaveFormat::Html, true),
            std::make_tuple(SaveFormat::Epub, true),
            std::make_tuple(SaveFormat::Mhtml, false),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportTextBoxAsSvgEpub, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportTextBoxAsSvgEpub(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportTextBoxAsSvgEpub, ::testing::ValuesIn(ExHtmlSaveOptions_ExportTextBoxAsSvgEpub::TestCases()));

using ExHtmlSaveOptions_ControlListLabelsExport_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ControlListLabelsExport)>::type;

struct ExHtmlSaveOptions_ControlListLabelsExport : public ExHtmlSaveOptions,
                                                   public ApiExamples::ExHtmlSaveOptions,
                                                   public ::testing::WithParamInterface<ExHtmlSaveOptions_ControlListLabelsExport_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(ExportListLabels::Auto),
            std::make_tuple(ExportListLabels::AsInlineText),
            std::make_tuple(ExportListLabels::ByHtmlTags),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ControlListLabelsExport, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ControlListLabelsExport(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ControlListLabelsExport, ::testing::ValuesIn(ExHtmlSaveOptions_ControlListLabelsExport::TestCases()));

using ExHtmlSaveOptions_ExportUrlForLinkedImage_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportUrlForLinkedImage)>::type;

struct ExHtmlSaveOptions_ExportUrlForLinkedImage : public ExHtmlSaveOptions,
                                                   public ApiExamples::ExHtmlSaveOptions,
                                                   public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportUrlForLinkedImage_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportUrlForLinkedImage, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportUrlForLinkedImage(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportUrlForLinkedImage, ::testing::ValuesIn(ExHtmlSaveOptions_ExportUrlForLinkedImage::TestCases()));

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

using ExHtmlSaveOptions_Html5Support_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::Html5Support)>::type;

struct ExHtmlSaveOptions_Html5Support : public ExHtmlSaveOptions,
                                        public ApiExamples::ExHtmlSaveOptions,
                                        public ::testing::WithParamInterface<ExHtmlSaveOptions_Html5Support_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(HtmlVersion::Html5),
            std::make_tuple(HtmlVersion::Xhtml),
        };
    }
};

TEST_P(ExHtmlSaveOptions_Html5Support, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Html5Support(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_Html5Support, ::testing::ValuesIn(ExHtmlSaveOptions_Html5Support::TestCases()));

using ExHtmlSaveOptions_ExportFonts_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportFonts)>::type;

struct ExHtmlSaveOptions_ExportFonts : public ExHtmlSaveOptions,
                                       public ApiExamples::ExHtmlSaveOptions,
                                       public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportFonts_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportFonts, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportFonts(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportFonts, ::testing::ValuesIn(ExHtmlSaveOptions_ExportFonts::TestCases()));

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

using ExHtmlSaveOptions_NegativeIndent_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::NegativeIndent)>::type;

struct ExHtmlSaveOptions_NegativeIndent : public ExHtmlSaveOptions,
                                          public ApiExamples::ExHtmlSaveOptions,
                                          public ::testing::WithParamInterface<ExHtmlSaveOptions_NegativeIndent_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_NegativeIndent, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->NegativeIndent(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_NegativeIndent, ::testing::ValuesIn(ExHtmlSaveOptions_NegativeIndent::TestCases()));

TEST_F(ExHtmlSaveOptions, FolderAlias)
{
    s_instance->FolderAlias();
}

TEST_F(ExHtmlSaveOptions, SaveExportedFonts)
{
    s_instance->SaveExportedFonts();
}

using ExHtmlSaveOptions_HtmlVersions_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::HtmlVersions)>::type;

struct ExHtmlSaveOptions_HtmlVersions : public ExHtmlSaveOptions,
                                        public ApiExamples::ExHtmlSaveOptions,
                                        public ::testing::WithParamInterface<ExHtmlSaveOptions_HtmlVersions_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(HtmlVersion::Html5),
            std::make_tuple(HtmlVersion::Xhtml),
        };
    }
};

TEST_P(ExHtmlSaveOptions_HtmlVersions, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->HtmlVersions(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_HtmlVersions, ::testing::ValuesIn(ExHtmlSaveOptions_HtmlVersions::TestCases()));

using ExHtmlSaveOptions_ExportXhtmlTransitional_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportXhtmlTransitional)>::type;

struct ExHtmlSaveOptions_ExportXhtmlTransitional : public ExHtmlSaveOptions,
                                                   public ApiExamples::ExHtmlSaveOptions,
                                                   public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportXhtmlTransitional_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportXhtmlTransitional, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportXhtmlTransitional(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportXhtmlTransitional, ::testing::ValuesIn(ExHtmlSaveOptions_ExportXhtmlTransitional::TestCases()));

TEST_F(ExHtmlSaveOptions, EpubHeadings)
{
    s_instance->EpubHeadings();
}

TEST_F(ExHtmlSaveOptions, Doc2EpubSaveOptions)
{
    s_instance->Doc2EpubSaveOptions();
}

using ExHtmlSaveOptions_ContentIdUrls_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ContentIdUrls)>::type;

struct ExHtmlSaveOptions_ContentIdUrls : public ExHtmlSaveOptions,
                                         public ApiExamples::ExHtmlSaveOptions,
                                         public ::testing::WithParamInterface<ExHtmlSaveOptions_ContentIdUrls_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ContentIdUrls, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ContentIdUrls(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ContentIdUrls, ::testing::ValuesIn(ExHtmlSaveOptions_ContentIdUrls::TestCases()));

using ExHtmlSaveOptions_DropDownFormField_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::DropDownFormField)>::type;

struct ExHtmlSaveOptions_DropDownFormField : public ExHtmlSaveOptions,
                                             public ApiExamples::ExHtmlSaveOptions,
                                             public ::testing::WithParamInterface<ExHtmlSaveOptions_DropDownFormField_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_DropDownFormField, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DropDownFormField(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_DropDownFormField, ::testing::ValuesIn(ExHtmlSaveOptions_DropDownFormField::TestCases()));

using ExHtmlSaveOptions_ExportImagesAsBase64_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportImagesAsBase64)>::type;

struct ExHtmlSaveOptions_ExportImagesAsBase64 : public ExHtmlSaveOptions,
                                                public ApiExamples::ExHtmlSaveOptions,
                                                public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportImagesAsBase64_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportImagesAsBase64, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportImagesAsBase64(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportImagesAsBase64, ::testing::ValuesIn(ExHtmlSaveOptions_ExportImagesAsBase64::TestCases()));

TEST_F(ExHtmlSaveOptions, ExportFontsAsBase64)
{
    s_instance->ExportFontsAsBase64();
}

using ExHtmlSaveOptions_ExportLanguageInformation_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportLanguageInformation)>::type;

struct ExHtmlSaveOptions_ExportLanguageInformation : public ExHtmlSaveOptions,
                                                     public ApiExamples::ExHtmlSaveOptions,
                                                     public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportLanguageInformation_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportLanguageInformation, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportLanguageInformation(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportLanguageInformation, ::testing::ValuesIn(ExHtmlSaveOptions_ExportLanguageInformation::TestCases()));

using ExHtmlSaveOptions_List_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::List)>::type;

struct ExHtmlSaveOptions_List : public ExHtmlSaveOptions,
                                public ApiExamples::ExHtmlSaveOptions,
                                public ::testing::WithParamInterface<ExHtmlSaveOptions_List_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(ExportListLabels::AsInlineText),
            std::make_tuple(ExportListLabels::Auto),
            std::make_tuple(ExportListLabels::ByHtmlTags),
        };
    }
};

TEST_P(ExHtmlSaveOptions_List, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->List(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_List, ::testing::ValuesIn(ExHtmlSaveOptions_List::TestCases()));

using ExHtmlSaveOptions_ExportPageMargins_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportPageMargins)>::type;

struct ExHtmlSaveOptions_ExportPageMargins : public ExHtmlSaveOptions,
                                             public ApiExamples::ExHtmlSaveOptions,
                                             public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportPageMargins_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportPageMargins, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportPageMargins(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportPageMargins, ::testing::ValuesIn(ExHtmlSaveOptions_ExportPageMargins::TestCases()));

using ExHtmlSaveOptions_ExportPageSetup_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportPageSetup)>::type;

struct ExHtmlSaveOptions_ExportPageSetup : public ExHtmlSaveOptions,
                                           public ApiExamples::ExHtmlSaveOptions,
                                           public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportPageSetup_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportPageSetup, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportPageSetup(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportPageSetup, ::testing::ValuesIn(ExHtmlSaveOptions_ExportPageSetup::TestCases()));

using ExHtmlSaveOptions_RelativeFontSize_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::RelativeFontSize)>::type;

struct ExHtmlSaveOptions_RelativeFontSize : public ExHtmlSaveOptions,
                                            public ApiExamples::ExHtmlSaveOptions,
                                            public ::testing::WithParamInterface<ExHtmlSaveOptions_RelativeFontSize_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_RelativeFontSize, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RelativeFontSize(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_RelativeFontSize, ::testing::ValuesIn(ExHtmlSaveOptions_RelativeFontSize::TestCases()));

using ExHtmlSaveOptions_ExportTextBox_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportTextBox)>::type;

struct ExHtmlSaveOptions_ExportTextBox : public ExHtmlSaveOptions,
                                         public ApiExamples::ExHtmlSaveOptions,
                                         public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportTextBox_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportTextBox, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportTextBox(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportTextBox, ::testing::ValuesIn(ExHtmlSaveOptions_ExportTextBox::TestCases()));

using ExHtmlSaveOptions_RoundTripInformation_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::RoundTripInformation)>::type;

struct ExHtmlSaveOptions_RoundTripInformation : public ExHtmlSaveOptions,
                                                public ApiExamples::ExHtmlSaveOptions,
                                                public ::testing::WithParamInterface<ExHtmlSaveOptions_RoundTripInformation_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_RoundTripInformation, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RoundTripInformation(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_RoundTripInformation, ::testing::ValuesIn(ExHtmlSaveOptions_RoundTripInformation::TestCases()));

using ExHtmlSaveOptions_ExportTocPageNumbers_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ExportTocPageNumbers)>::type;

struct ExHtmlSaveOptions_ExportTocPageNumbers : public ExHtmlSaveOptions,
                                                public ApiExamples::ExHtmlSaveOptions,
                                                public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportTocPageNumbers_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportTocPageNumbers, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportTocPageNumbers(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportTocPageNumbers, ::testing::ValuesIn(ExHtmlSaveOptions_ExportTocPageNumbers::TestCases()));

using ExHtmlSaveOptions_FontSubsetting_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::FontSubsetting)>::type;

struct ExHtmlSaveOptions_FontSubsetting : public ExHtmlSaveOptions,
                                          public ApiExamples::ExHtmlSaveOptions,
                                          public ::testing::WithParamInterface<ExHtmlSaveOptions_FontSubsetting_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(0),
            std::make_tuple(1000000),
            std::make_tuple(std::numeric_limits<int32_t>::max()),
        };
    }
};

TEST_P(ExHtmlSaveOptions_FontSubsetting, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FontSubsetting(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_FontSubsetting, ::testing::ValuesIn(ExHtmlSaveOptions_FontSubsetting::TestCases()));

using ExHtmlSaveOptions_MetafileFormat_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::MetafileFormat)>::type;

struct ExHtmlSaveOptions_MetafileFormat : public ExHtmlSaveOptions,
                                          public ApiExamples::ExHtmlSaveOptions,
                                          public ::testing::WithParamInterface<ExHtmlSaveOptions_MetafileFormat_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(HtmlMetafileFormat::Png),
            std::make_tuple(HtmlMetafileFormat::Svg),
            std::make_tuple(HtmlMetafileFormat::EmfOrWmf),
        };
    }
};

TEST_P(ExHtmlSaveOptions_MetafileFormat, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MetafileFormat(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_MetafileFormat, ::testing::ValuesIn(ExHtmlSaveOptions_MetafileFormat::TestCases()));

using ExHtmlSaveOptions_OfficeMathOutputMode_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::OfficeMathOutputMode)>::type;

struct ExHtmlSaveOptions_OfficeMathOutputMode : public ExHtmlSaveOptions,
                                                public ApiExamples::ExHtmlSaveOptions,
                                                public ::testing::WithParamInterface<ExHtmlSaveOptions_OfficeMathOutputMode_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(HtmlOfficeMathOutputMode::Image),
            std::make_tuple(HtmlOfficeMathOutputMode::MathML),
            std::make_tuple(HtmlOfficeMathOutputMode::Text),
        };
    }
};

TEST_P(ExHtmlSaveOptions_OfficeMathOutputMode, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->OfficeMathOutputMode(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_OfficeMathOutputMode, ::testing::ValuesIn(ExHtmlSaveOptions_OfficeMathOutputMode::TestCases()));

using ExHtmlSaveOptions_ScaleImageToShapeSize_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::ScaleImageToShapeSize)>::type;

struct ExHtmlSaveOptions_ScaleImageToShapeSize : public ExHtmlSaveOptions,
                                                 public ApiExamples::ExHtmlSaveOptions,
                                                 public ::testing::WithParamInterface<ExHtmlSaveOptions_ScaleImageToShapeSize_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ScaleImageToShapeSize, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ScaleImageToShapeSize(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ScaleImageToShapeSize, ::testing::ValuesIn(ExHtmlSaveOptions_ScaleImageToShapeSize::TestCases()));

TEST_F(ExHtmlSaveOptions, ImageFolder)
{
    s_instance->ImageFolder();
}

TEST_F(ExHtmlSaveOptions, ImageSavingCallback)
{
    s_instance->ImageSavingCallback();
}

using ExHtmlSaveOptions_PrettyFormat_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlSaveOptions::PrettyFormat)>::type;

struct ExHtmlSaveOptions_PrettyFormat : public ExHtmlSaveOptions,
                                        public ApiExamples::ExHtmlSaveOptions,
                                        public ::testing::WithParamInterface<ExHtmlSaveOptions_PrettyFormat_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExHtmlSaveOptions_PrettyFormat, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PrettyFormat(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_PrettyFormat, ::testing::ValuesIn(ExHtmlSaveOptions_PrettyFormat::TestCases()));

}} // namespace ApiExamples::gtest_test
