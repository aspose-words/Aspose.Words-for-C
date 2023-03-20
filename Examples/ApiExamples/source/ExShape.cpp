// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExShape.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Ole;
using namespace Aspose::Words::Math;
using namespace Aspose::Words::Rendering;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
using namespace Aspose::Words::Tables;
namespace ApiExamples { namespace gtest_test {

class ExShape : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExShape> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExShape>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExShape> ExShape::s_instance;

using ExShape_Font_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::Font)>::type;

struct ExShape_Font : public ExShape, public ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_Font_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExShape_Font, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Font(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_Font, ::testing::ValuesIn(ExShape_Font::TestCases()));

TEST_F(ExShape, Rotate)
{
    s_instance->Rotate();
}

TEST_F(ExShape, AspectRatioLockedDefaultValue)
{
    s_instance->AspectRatioLockedDefaultValue();
}

TEST_F(ExShape, Coordinates)
{
    s_instance->Coordinates();
}

TEST_F(ExShape, GroupShape)
{
    s_instance->GroupShape_();
}

TEST_F(ExShape, IsTopLevel)
{
    s_instance->IsTopLevel();
}

TEST_F(ExShape, LocalToParent)
{
    s_instance->LocalToParent();
}

using ExShape_AnchorLocked_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::AnchorLocked)>::type;

struct ExShape_AnchorLocked : public ExShape, public ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_AnchorLocked_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExShape_AnchorLocked, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AnchorLocked(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_AnchorLocked, ::testing::ValuesIn(ExShape_AnchorLocked::TestCases()));

TEST_F(ExShape, DeleteAllShapes)
{
    s_instance->DeleteAllShapes();
}

TEST_F(ExShape, IsInline)
{
    s_instance->IsInline();
}

TEST_F(ExShape, Bounds)
{
    s_instance->Bounds();
}

TEST_F(ExShape, FlipShapeOrientation)
{
    s_instance->FlipShapeOrientation();
}

TEST_F(ExShape, TextureFill)
{
    s_instance->TextureFill();
}

TEST_F(ExShape, GradientFill)
{
    s_instance->GradientFill();
}

TEST_F(ExShape, GradientStops)
{
    s_instance->GradientStops();
}

TEST_F(ExShape, FillPattern)
{
    s_instance->FillPattern();
}

TEST_F(ExShape, Title)
{
    s_instance->Title();
}

TEST_F(ExShape, ReplaceTextboxesWithImages)
{
    s_instance->ReplaceTextboxesWithImages();
}

TEST_F(ExShape, CreateTextBox)
{
    s_instance->CreateTextBox();
}

TEST_F(ExShape, ZOrder)
{
    s_instance->ZOrder();
}

TEST_F(ExShape, GetOleObjectRawData)
{
    s_instance->GetOleObjectRawData();
}

TEST_F(ExShape, OleControl)
{
    s_instance->OleControl_();
}

TEST_F(ExShape, OleLinks)
{
    s_instance->OleLinks();
}

TEST_F(ExShape, OleControlCollection)
{
    s_instance->OleControlCollection();
}

TEST_F(ExShape, SuggestedFileName)
{
    s_instance->SuggestedFileName();
}

TEST_F(ExShape, ObjectDidNotHaveSuggestedFileName)
{
    s_instance->ObjectDidNotHaveSuggestedFileName();
}

TEST_F(ExShape, ResolutionDefaultValues)
{
    s_instance->ResolutionDefaultValues();
}

TEST_F(ExShape, RenderOfficeMath)
{
    s_instance->RenderOfficeMath();
}

TEST_F(ExShape, OfficeMathDisplayException)
{
    s_instance->OfficeMathDisplayException();
}

TEST_F(ExShape, OfficeMathDefaultValue)
{
    s_instance->OfficeMathDefaultValue();
}

TEST_F(ExShape, OfficeMath)
{
    s_instance->OfficeMath_();
}

TEST_F(ExShape, CannotBeSetDisplayWithInlineJustification)
{
    s_instance->CannotBeSetDisplayWithInlineJustification();
}

TEST_F(ExShape, CannotBeSetInlineDisplayWithJustification)
{
    s_instance->CannotBeSetInlineDisplayWithJustification();
}

TEST_F(ExShape, OfficeMathDisplayNestedObjects)
{
    s_instance->OfficeMathDisplayNestedObjects();
}

using ExShape_WorkWithMathObjectType_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::WorkWithMathObjectType)>::type;

struct ExShape_WorkWithMathObjectType : public ExShape, public ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_WorkWithMathObjectType_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(0, MathObjectType::OMathPara), std::make_tuple(1, MathObjectType::OMath),           std::make_tuple(2, MathObjectType::Supercript),
            std::make_tuple(3, MathObjectType::Argument),  std::make_tuple(4, MathObjectType::SuperscriptPart),
        };
    }
};

TEST_P(ExShape_WorkWithMathObjectType, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->WorkWithMathObjectType(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_WorkWithMathObjectType, ::testing::ValuesIn(ExShape_WorkWithMathObjectType::TestCases()));

using ExShape_AspectRatio_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::AspectRatio)>::type;

struct ExShape_AspectRatio : public ExShape, public ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_AspectRatio_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExShape_AspectRatio, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AspectRatio(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_AspectRatio, ::testing::ValuesIn(ExShape_AspectRatio::TestCases()));

TEST_F(ExShape, MarkupLanguageByDefault)
{
    s_instance->MarkupLanguageByDefault();
}

using ExShape_MarkupLunguageForDifferentMsWordVersions_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::MarkupLunguageForDifferentMsWordVersions)>::type;

struct ExShape_MarkupLunguageForDifferentMsWordVersions : public ExShape,
                                                          public ApiExamples::ExShape,
                                                          public ::testing::WithParamInterface<ExShape_MarkupLunguageForDifferentMsWordVersions_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(MsWordVersion::Word2000, ShapeMarkupLanguage::Vml), std::make_tuple(MsWordVersion::Word2002, ShapeMarkupLanguage::Vml),
            std::make_tuple(MsWordVersion::Word2003, ShapeMarkupLanguage::Vml), std::make_tuple(MsWordVersion::Word2007, ShapeMarkupLanguage::Vml),
            std::make_tuple(MsWordVersion::Word2010, ShapeMarkupLanguage::Dml), std::make_tuple(MsWordVersion::Word2013, ShapeMarkupLanguage::Dml),
            std::make_tuple(MsWordVersion::Word2016, ShapeMarkupLanguage::Dml),
        };
    }
};

TEST_P(ExShape_MarkupLunguageForDifferentMsWordVersions, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MarkupLunguageForDifferentMsWordVersions(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_MarkupLunguageForDifferentMsWordVersions,
                         ::testing::ValuesIn(ExShape_MarkupLunguageForDifferentMsWordVersions::TestCases()));

TEST_F(ExShape, Stroke)
{
    s_instance->Stroke_();
}

TEST_F(ExShape, InsertOleObjectAsHtmlFile)
{
    s_instance->InsertOleObjectAsHtmlFile();
}

TEST_F(ExShape, InsertOlePackage)
{
    s_instance->InsertOlePackage();
}

TEST_F(ExShape, GetAccessToOlePackage)
{
    s_instance->GetAccessToOlePackage();
}

TEST_F(ExShape, Resize)
{
    s_instance->Resize();
}

TEST_F(ExShape, Calendar)
{
    s_instance->Calendar();
}

using ExShape_IsLayoutInCell_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::IsLayoutInCell)>::type;

struct ExShape_IsLayoutInCell : public ExShape, public ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_IsLayoutInCell_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExShape_IsLayoutInCell, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IsLayoutInCell(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_IsLayoutInCell, ::testing::ValuesIn(ExShape_IsLayoutInCell::TestCases()));

TEST_F(ExShape, ShapeInsertion)
{
    s_instance->ShapeInsertion();
}

TEST_F(ExShape, VisitShapes)
{
    s_instance->VisitShapes();
}

TEST_F(ExShape, SignatureLine)
{
    s_instance->SignatureLine_();
}

TEST_F(ExShape, TextBoxFitShapeToText)
{
    s_instance->TextBoxFitShapeToText();
}

TEST_F(ExShape, TextBoxMargins)
{
    s_instance->TextBoxMargins();
}

using ExShape_TextBoxContentsWrapMode_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::TextBoxContentsWrapMode)>::type;

struct ExShape_TextBoxContentsWrapMode : public ExShape, public ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_TextBoxContentsWrapMode_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(TextBoxWrapMode::None),
            std::make_tuple(TextBoxWrapMode::Square),
        };
    }
};

TEST_P(ExShape_TextBoxContentsWrapMode, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TextBoxContentsWrapMode(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_TextBoxContentsWrapMode, ::testing::ValuesIn(ExShape_TextBoxContentsWrapMode::TestCases()));

TEST_F(ExShape, TextBoxShapeType)
{
    s_instance->TextBoxShapeType();
}

TEST_F(ExShape, CreateLinkBetweenTextBoxes)
{
    s_instance->CreateLinkBetweenTextBoxes();
}

using ExShape_VerticalAnchor_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::VerticalAnchor)>::type;

struct ExShape_VerticalAnchor : public ExShape, public ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_VerticalAnchor_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(TextBoxAnchor::Top),
            std::make_tuple(TextBoxAnchor::Middle),
            std::make_tuple(TextBoxAnchor::Bottom),
        };
    }
};

TEST_P(ExShape_VerticalAnchor, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->VerticalAnchor(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_VerticalAnchor, ::testing::ValuesIn(ExShape_VerticalAnchor::TestCases()));

TEST_F(ExShape, InsertTextPaths)
{
    s_instance->InsertTextPaths();
}

TEST_F(ExShape, ShapeRevision)
{
    s_instance->ShapeRevision();
}

TEST_F(ExShape, MoveRevisions)
{
    s_instance->MoveRevisions();
}

TEST_F(ExShape, AdjustWithEffects)
{
    s_instance->AdjustWithEffects();
}

TEST_F(ExShape, RenderAllShapes)
{
    s_instance->RenderAllShapes();
}

TEST_F(ExShape, DocumentHasSmartArtObject)
{
    s_instance->DocumentHasSmartArtObject();
}

TEST_F(ExShape, SkipMono_OfficeMathRenderer)
{
    RecordProperty("category", "SkipMono");

    s_instance->OfficeMathRenderer_();
}

TEST_F(ExShape, ShapeTypes)
{
    s_instance->ShapeTypes();
}

TEST_F(ExShape, IsDecorative)
{
    s_instance->IsDecorative();
}

}} // namespace ApiExamples::gtest_test
