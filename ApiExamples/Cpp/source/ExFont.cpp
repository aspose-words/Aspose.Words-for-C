// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExFont.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Text/EmphasisMark.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/SpecialChar.h>
#include <system/io/stream.h>
#include <system/shared_ptr.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Tables;
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

TEST_F(ExFont, FontInfoCollection)
{
    s_instance->FontInfoCollection_();
}

using WorkWithEmbeddedFonts_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExFont::WorkWithEmbeddedFonts)>::type;

struct WorkWithEmbeddedFontsVP : public ExFont, public ApiExamples::ExFont, public ::testing::WithParamInterface<WorkWithEmbeddedFonts_Args>
{
    static std::vector<WorkWithEmbeddedFonts_Args> TestCases()
    {
        return {
            std::make_tuple(true, false, false), std::make_tuple(true, true, false),   std::make_tuple(true, true, true),
            std::make_tuple(true, false, true),  std::make_tuple(false, false, false),
        };
    }
};

TEST_P(WorkWithEmbeddedFontsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->WorkWithEmbeddedFonts(get<0>(params), get<1>(params), get<2>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExFont, WorkWithEmbeddedFontsVP, ::testing::ValuesIn(WorkWithEmbeddedFontsVP::TestCases()));

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

TEST_F(ExFont, Names)
{
    s_instance->Names();
}

TEST_F(ExFont, ChangeStyle)
{
    s_instance->ChangeStyle();
}

TEST_F(ExFont, Style)
{
    s_instance->Style_();
}

TEST_F(ExFont, SubstitutionNotification)
{
    s_instance->SubstitutionNotification();
}

TEST_F(ExFont, GetAvailableFonts)
{
    s_instance->GetAvailableFonts();
}

TEST_F(ExFont, EnableFontSubstitution)
{
    s_instance->EnableFontSubstitution();
}

TEST_F(ExFont, DisableFontSubstitution)
{
    s_instance->DisableFontSubstitution();
}

TEST_F(ExFont, SubstitutionWarnings)
{
    s_instance->SubstitutionWarnings();
}

TEST_F(ExFont, SubstitutionWarningsClosestMatch)
{
    s_instance->SubstitutionWarningsClosestMatch();
}

TEST_F(ExFont, SetFontAutoColor)
{
    s_instance->SetFontAutoColor();
}

TEST_F(ExFont, RemoveHiddenContentFromDocument)
{
    s_instance->RemoveHiddenContentFromDocument();
}

TEST_F(ExFont, BlankDocumentFonts)
{
    s_instance->BlankDocumentFonts();
}

TEST_F(ExFont, ExtractEmbeddedFont)
{
    s_instance->ExtractEmbeddedFont();
}

TEST_F(ExFont, GetFontInfoFromFile)
{
    s_instance->GetFontInfoFromFile();
}

TEST_F(ExFont, FontSourceFile)
{
    s_instance->FontSourceFile();
}

TEST_F(ExFont, FontSourceFolder)
{
    s_instance->FontSourceFolder();
}

TEST_F(ExFont, FontSourceMemory)
{
    s_instance->FontSourceMemory();
}

TEST_F(ExFont, FontSourceSystem)
{
    s_instance->FontSourceSystem();
}

TEST_F(ExFont, LoadFontFallbackSettingsFromFile)
{
    s_instance->LoadFontFallbackSettingsFromFile();
}

TEST_F(ExFont, LoadFontFallbackSettingsFromStream)
{
    s_instance->LoadFontFallbackSettingsFromStream();
}

TEST_F(ExFont, LoadNotoFontsFallbackSettings)
{
    s_instance->LoadNotoFontsFallbackSettings();
}

TEST_F(ExFont, DefaultFontSubstitutionRule)
{
    s_instance->DefaultFontSubstitutionRule_();
}

TEST_F(ExFont, FontConfigSubstitution)
{
    s_instance->FontConfigSubstitution();
}

TEST_F(ExFont, FallbackSettings)
{
    s_instance->FallbackSettings();
}

TEST_F(ExFont, FallbackSettingsCustom)
{
    s_instance->FallbackSettingsCustom();
}

TEST_F(ExFont, TableSubstitutionRule)
{
    s_instance->TableSubstitutionRule_();
}

TEST_F(ExFont, TableSubstitutionRuleCustom)
{
    s_instance->TableSubstitutionRuleCustom();
}

TEST_F(ExFont, ResolveFontsBeforeLoadingDocument)
{
    s_instance->ResolveFontsBeforeLoadingDocument();
}

TEST_F(ExFont, SkipMono_LineSpacing)
{
    RecordProperty("category", "SkipMono");

    s_instance->LineSpacing();
}

TEST_F(ExFont, HasDmlEffect)
{
    s_instance->HasDmlEffect();
}

TEST_F(ExFont, StreamFontSourceFileRendering)
{
    s_instance->StreamFontSourceFileRendering();
}

TEST_F(ExFont, SkipMono_IgnoreOnJenkins_CheckScanUserFontsFolder)
{
    RecordProperty("category", "IgnoreOnJenkins");
    RecordProperty("category1", "SkipMono");

    s_instance->CheckScanUserFontsFolder();
}

using SetEmphasisMark_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExFont::SetEmphasisMark)>::type;

struct SetEmphasisMarkVP : public ExFont, public ApiExamples::ExFont, public ::testing::WithParamInterface<SetEmphasisMark_Args>
{
    static std::vector<SetEmphasisMark_Args> TestCases()
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

TEST_P(SetEmphasisMarkVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SetEmphasisMark(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExFont, SetEmphasisMarkVP, ::testing::ValuesIn(SetEmphasisMarkVP::TestCases()));

}} // namespace ApiExamples::gtest_test
