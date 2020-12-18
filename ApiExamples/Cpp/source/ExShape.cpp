// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExShape.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeMarkupLanguage.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Math/MathObjectType.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <drawing/color.h>
#include <system/shared_ptr.h>
#include <system/string.h>
#include <system/text/string_builder.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Math;
using namespace Aspose::Words::Settings;
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

TEST_F(ExShape, Insert)
{
    s_instance->Insert();
}

TEST_F(ExShape, AspectRatioLockedDefaultValue)
{
    s_instance->AspectRatioLockedDefaultValue();
}

TEST_F(ExShape, Coordinates)
{
    s_instance->Coordinates();
}

TEST_F(ExShape, InsertGroupShape)
{
    s_instance->InsertGroupShape();
}

TEST_F(ExShape, DeleteAllShapes)
{
    s_instance->DeleteAllShapes();
}

TEST_F(ExShape, CheckShapeInline)
{
    s_instance->CheckShapeInline();
}

TEST_F(ExShape, LineFlipOrientation)
{
    s_instance->LineFlipOrientation();
}

TEST_F(ExShape, Fill)
{
    s_instance->Fill_();
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

TEST_F(ExShape, GetActiveXControlProperties)
{
    s_instance->GetActiveXControlProperties();
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

TEST_F(ExShape, SaveShapeObjectAsImage)
{
    s_instance->SaveShapeObjectAsImage();
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

using WorkWithMathObjectType_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::WorkWithMathObjectType)>::type;

struct WorkWithMathObjectTypeVP : public ExShape, public ApiExamples::ExShape, public ::testing::WithParamInterface<WorkWithMathObjectType_Args>
{
    static std::vector<WorkWithMathObjectType_Args> TestCases()
    {
        return {
            std::make_tuple(0, MathObjectType::OMathPara), std::make_tuple(1, MathObjectType::OMath),           std::make_tuple(2, MathObjectType::Supercript),
            std::make_tuple(3, MathObjectType::Argument),  std::make_tuple(4, MathObjectType::SuperscriptPart),
        };
    }
};

TEST_P(WorkWithMathObjectTypeVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->WorkWithMathObjectType(get<0>(params), get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExShape, WorkWithMathObjectTypeVP, ::testing::ValuesIn(WorkWithMathObjectTypeVP::TestCases()));

using AspectRatioLocked_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::AspectRatioLocked)>::type;

struct AspectRatioLockedVP : public ExShape, public ApiExamples::ExShape, public ::testing::WithParamInterface<AspectRatioLocked_Args>
{
    static std::vector<AspectRatioLocked_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(AspectRatioLockedVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AspectRatioLocked(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExShape, AspectRatioLockedVP, ::testing::ValuesIn(AspectRatioLockedVP::TestCases()));

TEST_F(ExShape, MarkupLunguageByDefault)
{
    s_instance->MarkupLunguageByDefault();
}

using MarkupLunguageForDifferentMsWordVersions_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::MarkupLunguageForDifferentMsWordVersions)>::type;

struct MarkupLunguageForDifferentMsWordVersionsVP : public ExShape,
                                                    public ApiExamples::ExShape,
                                                    public ::testing::WithParamInterface<MarkupLunguageForDifferentMsWordVersions_Args>
{
    static std::vector<MarkupLunguageForDifferentMsWordVersions_Args> TestCases()
    {
        return {
            std::make_tuple(MsWordVersion::Word2000, ShapeMarkupLanguage::Vml), std::make_tuple(MsWordVersion::Word2002, ShapeMarkupLanguage::Vml),
            std::make_tuple(MsWordVersion::Word2003, ShapeMarkupLanguage::Vml), std::make_tuple(MsWordVersion::Word2007, ShapeMarkupLanguage::Vml),
            std::make_tuple(MsWordVersion::Word2010, ShapeMarkupLanguage::Dml), std::make_tuple(MsWordVersion::Word2013, ShapeMarkupLanguage::Dml),
            std::make_tuple(MsWordVersion::Word2016, ShapeMarkupLanguage::Dml),
        };
    }
};

TEST_P(MarkupLunguageForDifferentMsWordVersionsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MarkupLunguageForDifferentMsWordVersions(get<0>(params), get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExShape, MarkupLunguageForDifferentMsWordVersionsVP, ::testing::ValuesIn(MarkupLunguageForDifferentMsWordVersionsVP::TestCases()));

TEST_F(ExShape, ChangeStrokeProperties)
{
    s_instance->ChangeStrokeProperties();
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

TEST_F(ExShape, LayoutInTableCell)
{
    s_instance->LayoutInTableCell();
}

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

TEST_F(ExShape, TextBox)
{
    s_instance->TextBox_();
}

TEST_F(ExShape, TextBoxShapeType)
{
    s_instance->TextBoxShapeType();
}

TEST_F(ExShape, CreateLinkBetweenTextBoxes)
{
    s_instance->CreateLinkBetweenTextBoxes();
}

TEST_F(ExShape, GetTextBoxAndChangeTextAnchor)
{
    s_instance->GetTextBoxAndChangeTextAnchor();
}

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

}} // namespace ApiExamples::gtest_test
