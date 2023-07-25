// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExFontSettings.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Loading;
namespace ApiExamples { namespace gtest_test {

class ExFontSettings : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExFontSettings> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExFontSettings>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExFontSettings> ExFontSettings::s_instance;

TEST_F(ExFontSettings, DefaultFontInstance)
{
    s_instance->DefaultFontInstance();
}

TEST_F(ExFontSettings, DefaultFontName)
{
    s_instance->DefaultFontName();
}

TEST_F(ExFontSettings, UpdatePageLayoutWarnings)
{
    s_instance->UpdatePageLayoutWarnings();
}

TEST_F(ExFontSettings, SubstitutionWarning)
{
    s_instance->SubstitutionWarning();
}

TEST_F(ExFontSettings, FontSourceWarning)
{
    s_instance->FontSourceWarning();
}

TEST_F(ExFontSettings, EnableFontSubstitution)
{
    s_instance->EnableFontSubstitution();
}

TEST_F(ExFontSettings, SubstitutionWarningsClosestMatch)
{
    s_instance->SubstitutionWarningsClosestMatch();
}

TEST_F(ExFontSettings, DisableFontSubstitution)
{
    s_instance->DisableFontSubstitution();
}

TEST_F(ExFontSettings, SkipMono_SubstitutionWarnings)
{
    RecordProperty("category", "SkipMono");

    s_instance->SubstitutionWarnings();
}

TEST_F(ExFontSettings, GetSubstitutionWithoutSuffixes)
{
    s_instance->GetSubstitutionWithoutSuffixes();
}

TEST_F(ExFontSettings, FontSourceFile)
{
    s_instance->FontSourceFile();
}

TEST_F(ExFontSettings, FontSourceFolder)
{
    s_instance->FontSourceFolder();
}

using ExFontSettings_SetFontsFolder_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExFontSettings::SetFontsFolder)>::type;

struct ExFontSettings_SetFontsFolder : public ExFontSettings,
                                       public ApiExamples::ExFontSettings,
                                       public ::testing::WithParamInterface<ExFontSettings_SetFontsFolder_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExFontSettings_SetFontsFolder, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SetFontsFolder(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExFontSettings_SetFontsFolder, ::testing::ValuesIn(ExFontSettings_SetFontsFolder::TestCases()));

using ExFontSettings_SetFontsFolders_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExFontSettings::SetFontsFolders)>::type;

struct ExFontSettings_SetFontsFolders : public ExFontSettings,
                                        public ApiExamples::ExFontSettings,
                                        public ::testing::WithParamInterface<ExFontSettings_SetFontsFolders_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExFontSettings_SetFontsFolders, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SetFontsFolders(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExFontSettings_SetFontsFolders, ::testing::ValuesIn(ExFontSettings_SetFontsFolders::TestCases()));

TEST_F(ExFontSettings, AddFontSource)
{
    s_instance->AddFontSource();
}

TEST_F(ExFontSettings, SetSpecifyFontFolder)
{
    s_instance->SetSpecifyFontFolder();
}

TEST_F(ExFontSettings, TableSubstitution)
{
    s_instance->TableSubstitution();
}

TEST_F(ExFontSettings, SetSpecifyFontFolders)
{
    s_instance->SetSpecifyFontFolders();
}

TEST_F(ExFontSettings, AddFontSubstitutes)
{
    s_instance->AddFontSubstitutes();
}

TEST_F(ExFontSettings, FontSourceMemory)
{
    s_instance->FontSourceMemory();
}

TEST_F(ExFontSettings, FontSourceSystem)
{
    s_instance->FontSourceSystem();
}

TEST_F(ExFontSettings, LoadFontFallbackSettingsFromFile)
{
    s_instance->LoadFontFallbackSettingsFromFile();
}

TEST_F(ExFontSettings, LoadFontFallbackSettingsFromStream)
{
    s_instance->LoadFontFallbackSettingsFromStream();
}

TEST_F(ExFontSettings, LoadNotoFontsFallbackSettings)
{
    s_instance->LoadNotoFontsFallbackSettings();
}

TEST_F(ExFontSettings, DefaultFontSubstitutionRule)
{
    s_instance->DefaultFontSubstitutionRule_();
}

TEST_F(ExFontSettings, FontConfigSubstitution)
{
    s_instance->FontConfigSubstitution();
}

TEST_F(ExFontSettings, FallbackSettingsCustom)
{
    s_instance->FallbackSettingsCustom();
}

TEST_F(ExFontSettings, TableSubstitutionRule)
{
    s_instance->TableSubstitutionRule_();
}

TEST_F(ExFontSettings, TableSubstitutionRuleCustom)
{
    s_instance->TableSubstitutionRuleCustom();
}

TEST_F(ExFontSettings, ResolveFontsBeforeLoadingDocument)
{
    s_instance->ResolveFontsBeforeLoadingDocument();
}

TEST_F(ExFontSettings, StreamFontSourceFileRendering)
{
    s_instance->StreamFontSourceFileRendering();
}

}} // namespace ApiExamples::gtest_test
