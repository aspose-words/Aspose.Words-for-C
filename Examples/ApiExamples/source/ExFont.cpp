// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExFont.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Themes;
namespace ApiExamples { namespace gtest_test {

class ExFont : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExFont> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExFont>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExFont> ExFont::s_instance;

TEST_F(ExFont, CreateFormattedRun)
{
    s_instance->CreateFormattedRun();
}

TEST_F(ExFont, Caps)
{
    s_instance->Caps();
}

TEST_F(ExFont, GetDocumentFonts)
{
    s_instance->GetDocumentFonts();
}

TEST_F(ExFont, DefaultValuesEmbeddedFontsParameters)
{
    s_instance->DefaultValuesEmbeddedFontsParameters();
}

using ExFont_FontInfoCollection_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExFont::FontInfoCollection)>::type;

struct ExFont_FontInfoCollection : public ExFont, public ApiExamples::ExFont, public ::testing::WithParamInterface<ExFont_FontInfoCollection_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExFont_FontInfoCollection, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FontInfoCollection(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExFont_FontInfoCollection, ::testing::ValuesIn(ExFont_FontInfoCollection::TestCases()));

using ExFont_WorkWithEmbeddedFonts_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExFont::WorkWithEmbeddedFonts)>::type;

struct ExFont_WorkWithEmbeddedFonts : public ExFont, public ApiExamples::ExFont, public ::testing::WithParamInterface<ExFont_WorkWithEmbeddedFonts_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true, false, false), std::make_tuple(true, true, false),   std::make_tuple(true, true, true),
            std::make_tuple(true, false, true),  std::make_tuple(false, false, false),
        };
    }
};

TEST_P(ExFont_WorkWithEmbeddedFonts, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->WorkWithEmbeddedFonts(std::get<0>(params), std::get<1>(params), std::get<2>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExFont_WorkWithEmbeddedFonts, ::testing::ValuesIn(ExFont_WorkWithEmbeddedFonts::TestCases()));

TEST_F(ExFont, StrikeThrough)
{
    s_instance->StrikeThrough();
}

TEST_F(ExFont, PositionSubscript)
{
    s_instance->PositionSubscript();
}

TEST_F(ExFont, ScalingSpacing)
{
    s_instance->ScalingSpacing();
}

TEST_F(ExFont, Italic)
{
    s_instance->Italic();
}

TEST_F(ExFont, EngraveEmboss)
{
    s_instance->EngraveEmboss();
}

TEST_F(ExFont, Shadow)
{
    s_instance->Shadow();
}

TEST_F(ExFont, Outline)
{
    s_instance->Outline();
}

TEST_F(ExFont, Hidden)
{
    s_instance->Hidden();
}

TEST_F(ExFont, Kerning)
{
    s_instance->Kerning();
}

TEST_F(ExFont, NoProofing)
{
    s_instance->NoProofing();
}

TEST_F(ExFont, LocaleId)
{
    s_instance->LocaleId();
}

TEST_F(ExFont, Underlines)
{
    s_instance->Underlines();
}

TEST_F(ExFont, ComplexScript)
{
    s_instance->ComplexScript();
}

TEST_F(ExFont, SparklingText)
{
    s_instance->SparklingText();
}

TEST_F(ExFont, Shading)
{
    s_instance->Shading_();
}

TEST_F(ExFont, SkipMono_Bidi)
{
    RecordProperty("category", "SkipMono");

    s_instance->Bidi();
}

TEST_F(ExFont, FarEast)
{
    s_instance->FarEast();
}

TEST_F(ExFont, NameAscii)
{
    s_instance->NameAscii();
}

TEST_F(ExFont, ChangeStyle)
{
    s_instance->ChangeStyle();
}

TEST_F(ExFont, BuiltIn)
{
    s_instance->BuiltIn();
}

TEST_F(ExFont, Style)
{
    s_instance->Style_();
}

TEST_F(ExFont, GetAvailableFonts)
{
    s_instance->GetAvailableFonts();
}

TEST_F(ExFont, SetFontAutoColor)
{
    s_instance->SetFontAutoColor();
}

TEST_F(ExFont, RemoveHiddenContentFromDocument)
{
    s_instance->RemoveHiddenContentFromDocument();
}

TEST_F(ExFont, DefaultFonts)
{
    s_instance->DefaultFonts();
}

TEST_F(ExFont, ExtractEmbeddedFont)
{
    s_instance->ExtractEmbeddedFont();
}

TEST_F(ExFont, GetFontInfoFromFile)
{
    s_instance->GetFontInfoFromFile();
}

TEST_F(ExFont, LineSpacing)
{
    s_instance->LineSpacing();
}

TEST_F(ExFont, HasDmlEffect)
{
    s_instance->HasDmlEffect();
}

TEST_F(ExFont, DISABLED_IgnoreOnJenkins_CheckScanUserFontsFolder)
{
    RecordProperty("category", "IgnoreOnJenkins");

    s_instance->CheckScanUserFontsFolder();
}

using ExFont_SetEmphasisMark_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExFont::SetEmphasisMark)>::type;

struct ExFont_SetEmphasisMark : public ExFont, public ApiExamples::ExFont, public ::testing::WithParamInterface<ExFont_SetEmphasisMark_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(EmphasisMark::None),
            std::make_tuple(EmphasisMark::OverComma),
            std::make_tuple(EmphasisMark::OverSolidCircle),
            std::make_tuple(EmphasisMark::OverWhiteCircle),
            std::make_tuple(EmphasisMark::UnderSolidCircle),
        };
    }
};

TEST_P(ExFont_SetEmphasisMark, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SetEmphasisMark(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExFont_SetEmphasisMark, ::testing::ValuesIn(ExFont_SetEmphasisMark::TestCases()));

TEST_F(ExFont, ThemeFontsColors)
{
    s_instance->ThemeFontsColors();
}

TEST_F(ExFont, CreateThemedStyle)
{
    s_instance->CreateThemedStyle();
}

}} // namespace ApiExamples::gtest_test
