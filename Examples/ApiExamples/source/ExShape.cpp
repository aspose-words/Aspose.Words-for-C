// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExShape.h"

#include <testing/test_predicates.h>
#include <system/type_info.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/path.h>
#include <system/io/memory_stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file_info.h>
#include <system/io/file.h>
#include <system/guid.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/enum_helpers.h>
#include <system/enum.h>
#include <system/details/dispose_guard.h>
#include <system/convert.h>
#include <system/collections/list.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/size_f.h>
#include <drawing/size.h>
#include <drawing/rectangle_f.h>
#include <drawing/rectangle.h>
#include <drawing/point_f.h>
#include <drawing/point.h>
#include <cstdint>
#include <Aspose.Words.Cpp/RW/Ole/ActiveX/Forms2/TextBoxControl.h>
#include <Aspose.Words.Cpp/RW/Ole/ActiveX/Forms2/OptionButtonControl.h>
#include <Aspose.Words.Cpp/RW/Ole/ActiveX/Forms2/CommandButtonControl.h>
#include <Aspose.Words.Cpp/RW/Ole/ActiveX/Forms2/CheckBoxControl.h>
#include <Aspose.Words.Cpp/Rendering/ShapeRenderer.h>
#include <Aspose.Words.Cpp/Rendering/OfficeMathRenderer.h>
#include <Aspose.Words.Cpp/Model/Themes/ThemeColor.h>
#include <Aspose.Words.Cpp/Model/Text/Underline.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/HeightRule.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Styles/TableStyle.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/StoryType.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Math/OfficeMathJustification.h>
#include <Aspose.Words.Cpp/Model/Math/OfficeMathDisplayType.h>
#include <Aspose.Words.Cpp/Model/Math/OfficeMath.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapSide.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextureAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextPathAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextPath.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBox.h>
#include <Aspose.Words.Cpp/Model/Drawing/Stroke.h>
#include <Aspose.Words.Cpp/Model/Drawing/SoftEdgeFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/SignatureLine.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeLineStyle.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShadowType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShadowFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalSize.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalSize.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/ReflectionFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/PresetTexture.h>
#include <Aspose.Words.Cpp/Model/Drawing/PatternType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OlePackage.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OleFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OleControl.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/Forms2OleControlType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/Forms2OleControlCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/Forms2OleControl.h>
#include <Aspose.Words.Cpp/Model/Drawing/JoinStyle.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/GradientVariant.h>
#include <Aspose.Words.Cpp/Model/Drawing/GradientStyle.h>
#include <Aspose.Words.Cpp/Model/Drawing/GradientStopCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/GradientStop.h>
#include <Aspose.Words.Cpp/Model/Drawing/GlowFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/FlipOrientation.h>
#include <Aspose.Words.Cpp/Model/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Model/Drawing/EndCap.h>
#include <Aspose.Words.Cpp/Model/Drawing/DashStyle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartTitle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Model/Drawing/AdjustmentCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Adjustment.h>
#include <Aspose.Words.Cpp/Model/Document/SignatureLineOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>

#include "TestUtil.h"
#include "DocumentHelper.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Ole;
using namespace Aspose::Words::Math;
using namespace Aspose::Words::Rendering;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Themes;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(653748774u, ::Aspose::Words::ApiExamples::ExShape::ShapeAppearancePrinter, ThisTypeBaseTypesInfo);

ExShape::ShapeAppearancePrinter::ShapeAppearancePrinter() : mShapesVisited(0), mTextIndentLevel(0)
{
    mShapesVisited = 0;
    mTextIndentLevel = 0;
    mStringBuilder = System::MakeObject<System::Text::StringBuilder>();
}

void ExShape::ShapeAppearancePrinter::AppendLine(System::String text)
{
    for (int32_t i = 0; i < mTextIndentLevel; i++)
    {
        mStringBuilder->Append(u'\t');
    }
    
    mStringBuilder->AppendLine(text);
}

System::String ExShape::ShapeAppearancePrinter::GetText()
{
    return System::String::Format(u"Shapes visited: {0}\n{1}", mShapesVisited, mStringBuilder);
}

Aspose::Words::VisitorAction ExShape::ShapeAppearancePrinter::VisitShapeStart(System::SharedPtr<Aspose::Words::Drawing::Shape> shape)
{
    AppendLine(System::String::Format(u"Shape found: {0}", shape->get_ShapeType()));
    
    mTextIndentLevel++;
    
    if (shape->get_HasChart())
    {
        AppendLine(System::String::Format(u"Has chart: {0}", shape->get_Chart()->get_Title()->get_Text()));
    }
    
    AppendLine(System::String::Format(u"Extrusion enabled: {0}", shape->get_ExtrusionEnabled()));
    AppendLine(System::String::Format(u"Shadow enabled: {0}", shape->get_ShadowEnabled()));
    AppendLine(System::String::Format(u"StoryType: {0}", shape->get_StoryType()));
    
    if (shape->get_Stroked())
    {
        ASPOSE_EXPECT_EQ(shape->get_Stroke()->get_Color(), shape->get_StrokeColor());
        AppendLine(System::String::Format(u"Stroke colors: {0}, {1}", shape->get_Stroke()->get_Color(), shape->get_Stroke()->get_Color2()));
        AppendLine(System::String::Format(u"Stroke weight: {0}", shape->get_StrokeWeight()));
    }
    
    if (shape->get_Filled())
    {
        AppendLine(System::String::Format(u"Filled: {0}", shape->get_FillColor()));
    }
    
    if (shape->get_OleFormat() != nullptr)
    {
        AppendLine(System::String::Format(u"Ole found of type: {0}", shape->get_OleFormat()->get_ProgId()));
    }
    
    if (shape->get_SignatureLine() != nullptr)
    {
        AppendLine(System::String::Format(u"Found signature line for: {0}, {1}", shape->get_SignatureLine()->get_Signer(), shape->get_SignatureLine()->get_SignerTitle()));
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExShape::ShapeAppearancePrinter::VisitShapeEnd(System::SharedPtr<Aspose::Words::Drawing::Shape> shape)
{
    mTextIndentLevel--;
    mShapesVisited++;
    AppendLine(System::String::Format(u"End of {0}", shape->get_ShapeType()));
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExShape::ShapeAppearancePrinter::VisitGroupShapeStart(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape)
{
    AppendLine(System::String::Format(u"Shape group found: {0}", groupShape->get_ShapeType()));
    mTextIndentLevel++;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExShape::ShapeAppearancePrinter::VisitGroupShapeEnd(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape)
{
    mTextIndentLevel--;
    AppendLine(System::String::Format(u"End of {0}", groupShape->get_ShapeType()));
    
    return Aspose::Words::VisitorAction::Continue;
}


RTTI_INFO_IMPL_HASH(3829644881u, ::Aspose::Words::ApiExamples::ExShape, ThisTypeBaseTypesInfo);

System::SharedPtr<Aspose::Words::Drawing::Shape> ExShape::AppendWordArt(System::SharedPtr<Aspose::Words::Document> doc, System::String text, System::String textFontFamily, double shapeWidth, double shapeHeight, System::Drawing::Color wordArtFill, System::Drawing::Color line, Aspose::Words::Drawing::ShapeType wordArtShapeType)
{
    // Create an inline Shape, which will serve as a container for our WordArt.
    // The shape can only be a valid WordArt shape if we assign a WordArt-designated ShapeType to it.
    // These types will have "WordArt object" in the description,
    // and their enumerator constant names will all start with "Text".
    auto shape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, wordArtShapeType);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Inline);
    shape->set_Width(shapeWidth);
    shape->set_Height(shapeHeight);
    shape->set_FillColor(wordArtFill);
    shape->set_StrokeColor(line);
    
    shape->get_TextPath()->set_Text(text);
    shape->get_TextPath()->set_FontFamily(textFontFamily);
    
    auto para = System::ExplicitCast<Aspose::Words::Paragraph>(doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(System::MakeObject<Aspose::Words::Paragraph>(doc)));
    para->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(shape);
    return shape;
}

void ExShape::TestInsertTextPaths(System::String filename)
{
    auto doc = System::MakeObject<Aspose::Words::Document>(filename);
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Drawing::Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToList();
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, System::String::Empty, 480, 24, 0.0, 0.0, shapes->idx_get(0));
    ASSERT_TRUE(shapes->idx_get(0)->get_TextPath()->get_Bold());
    ASSERT_TRUE(shapes->idx_get(0)->get_TextPath()->get_Italic());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, System::String::Empty, 150, 24, 0.0, 0.0, shapes->idx_get(1));
    ASSERT_TRUE(shapes->idx_get(1)->get_TextPath()->get_On());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, System::String::Empty, 150, 24, 0.0, 0.0, shapes->idx_get(2));
    ASSERT_FALSE(shapes->idx_get(2)->get_TextPath()->get_On());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, System::String::Empty, 90, 24, 0.0, 0.0, shapes->idx_get(3));
    ASSERT_TRUE(shapes->idx_get(3)->get_TextPath()->get_Kerning());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, System::String::Empty, 100, 24, 0.0, 0.0, shapes->idx_get(4));
    ASSERT_FALSE(shapes->idx_get(4)->get_TextPath()->get_Kerning());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextCascadeDown, System::String::Empty, 120, 24, 0.0, 0.0, shapes->idx_get(5));
    ASSERT_NEAR(0.1, shapes->idx_get(5)->get_TextPath()->get_Spacing(), 0.01);
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextWave, System::String::Empty, 200, 36, 0.0, 0.0, shapes->idx_get(6));
    ASSERT_TRUE(shapes->idx_get(6)->get_TextPath()->get_RotateLetters());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextSlantUp, System::String::Empty, 300, 24, 0.0, 0.0, shapes->idx_get(7));
    ASSERT_TRUE(shapes->idx_get(7)->get_TextPath()->get_SameLetterHeights());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, System::String::Empty, 160, 24, 0.0, 0.0, shapes->idx_get(8));
    ASSERT_TRUE(shapes->idx_get(8)->get_TextPath()->get_FitShape());
    ASPOSE_ASSERT_EQ(24.0, shapes->idx_get(8)->get_TextPath()->get_Size());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, System::String::Empty, 160, 24, 0.0, 0.0, shapes->idx_get(9));
    ASSERT_FALSE(shapes->idx_get(9)->get_TextPath()->get_FitShape());
    ASPOSE_ASSERT_EQ(24.0, shapes->idx_get(9)->get_TextPath()->get_Size());
    ASSERT_EQ(Aspose::Words::Drawing::TextPathAlignment::Right, shapes->idx_get(9)->get_TextPath()->get_TextPathAlignment());
}


namespace gtest_test
{

class ExShape : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExShape> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExShape>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExShape> ExShape::s_instance;

} // namespace gtest_test

void ExShape::AltText()
{
    //ExStart
    //ExFor:ShapeBase.AlternativeText
    //ExFor:ShapeBase.Name
    //ExSummary:Shows how to use a shape's alternative text.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Cube, 150, 150);
    shape->set_Name(u"MyCube");
    
    shape->set_AlternativeText(u"Alt text for MyCube.");
    
    // We can access the alternative text of a shape by right-clicking it, and then via "Format AutoShape" -> "Alt Text".
    doc->Save(get_ArtifactsDir() + u"Shape.AltText.docx");
    
    // Save the document to HTML, and then delete the linked image that belongs to our shape.
    // The browser that is reading our HTML will display the alt text in place of the missing image.
    doc->Save(get_ArtifactsDir() + u"Shape.AltText.html");
    ASSERT_TRUE(System::IO::File::Exists(get_ArtifactsDir() + u"Shape.AltText.001.png"));
    //ExSkip
    System::IO::File::Delete(get_ArtifactsDir() + u"Shape.AltText.001.png");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.AltText.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Cube, u"MyCube", 150.0, 150.0, 0, 0, shape);
    ASSERT_EQ(u"Alt text for MyCube.", shape->get_AlternativeText());
    ASSERT_EQ(u"Times New Roman", shape->get_Font()->get_Name());
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.AltText.html");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Image, System::String::Empty, 151.5, 151.5, 0, 0, shape);
    ASSERT_EQ(u"Alt text for MyCube.", shape->get_AlternativeText());
    
    Aspose::Words::ApiExamples::TestUtil::FileContainsString(System::String(u"<img src=\"Shape.AltText.001.png\" width=\"202\" height=\"202\" alt=\"Alt text for MyCube.\" ") + u"style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />", get_ArtifactsDir() + u"Shape.AltText.html");
}

namespace gtest_test
{

TEST_F(ExShape, AltText)
{
    s_instance->AltText();
}

} // namespace gtest_test

void ExShape::Font(bool hideShape)
{
    //ExStart
    //ExFor:ShapeBase.Font
    //ExFor:ShapeBase.ParentParagraph
    //ExSummary:Shows how to insert a text box, and set the font of its contents.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Hello world!");
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 300, 50);
    builder->MoveTo(shape->get_LastParagraph());
    builder->Write(u"This text is inside the text box.");
    
    // Set the "Hidden" property of the shape's "Font" object to "true" to hide the text box from sight
    // and collapse the space that it would normally occupy.
    // Set the "Hidden" property of the shape's "Font" object to "false" to leave the text box visible.
    shape->get_Font()->set_Hidden(hideShape);
    
    // If the shape is visible, we will modify its appearance via the font object.
    if (!hideShape)
    {
        shape->get_Font()->set_HighlightColor(System::Drawing::Color::get_LightGray());
        shape->get_Font()->set_Color(System::Drawing::Color::get_Red());
        shape->get_Font()->set_Underline(Aspose::Words::Underline::Dash);
    }
    
    // Move the builder out of the text box back into the main document.
    builder->MoveTo(shape->get_ParentParagraph());
    
    builder->Writeln(u"\nThis text is outside the text box.");
    
    doc->Save(get_ArtifactsDir() + u"Shape.Font.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.Font.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASPOSE_ASSERT_EQ(hideShape, shape->get_Font()->get_Hidden());
    
    if (hideShape)
    {
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), shape->get_Font()->get_HighlightColor().ToArgb());
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), shape->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(Aspose::Words::Underline::None, shape->get_Font()->get_Underline());
    }
    else
    {
        ASSERT_EQ(System::Drawing::Color::get_Silver().ToArgb(), shape->get_Font()->get_HighlightColor().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), shape->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(Aspose::Words::Underline::Dash, shape->get_Font()->get_Underline());
    }
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100002", 300.0, 50.0, 0, 0, shape);
    ASSERT_EQ(u"This text is inside the text box.", shape->GetText().Trim());
    ASSERT_EQ(u"Hello world!\rThis text is inside the text box.\r\rThis text is outside the text box.", doc->GetText().Trim());
}

namespace gtest_test
{

using ExShape_Font_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExShape::Font)>::type;

struct ExShape_Font : public ExShape, public Aspose::Words::ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_Font_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
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

} // namespace gtest_test

void ExShape::Rotate()
{
    //ExStart
    //ExFor:ShapeBase.CanHaveImage
    //ExFor:ShapeBase.Rotation
    //ExSummary:Shows how to insert and rotate an image.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a shape with an image.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    ASSERT_TRUE(shape->get_CanHaveImage());
    ASSERT_TRUE(shape->get_HasImage());
    
    // Rotate the image 45 degrees clockwise.
    shape->set_Rotation(45);
    
    doc->Save(get_ArtifactsDir() + u"Shape.Rotate.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.Rotate.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Image, System::String::Empty, 300.0, 300.0, 0, 0, shape);
    ASSERT_TRUE(shape->get_CanHaveImage());
    ASSERT_TRUE(shape->get_HasImage());
    ASPOSE_ASSERT_EQ(45.0, shape->get_Rotation());
}

namespace gtest_test
{

TEST_F(ExShape, Rotate)
{
    s_instance->Rotate();
}

} // namespace gtest_test

void ExShape::Coordinates()
{
    //ExStart
    //ExFor:ShapeBase.DistanceBottom
    //ExFor:ShapeBase.DistanceLeft
    //ExFor:ShapeBase.DistanceRight
    //ExFor:ShapeBase.DistanceTop
    //ExSummary:Shows how to set the wrapping distance for a text that surrounds a shape.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a rectangle and, get the text to wrap tightly around its bounds.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 150, 150);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Tight);
    
    // Set the minimum distance between the shape and surrounding text to 40pt from all sides.
    shape->set_DistanceTop(40);
    shape->set_DistanceBottom(40);
    shape->set_DistanceLeft(40);
    shape->set_DistanceRight(40);
    
    // Move the shape closer to the center of the page, and then rotate the shape 60 degrees clockwise.
    shape->set_Top(75);
    shape->set_Left(150);
    shape->set_Rotation(60);
    
    // Add text that will wrap around the shape.
    builder->get_Font()->set_Size(24);
    builder->Write(System::String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ") + u"Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");
    
    doc->Save(get_ArtifactsDir() + u"Shape.Coordinates.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.Coordinates.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, u"Rectangle 100002", 150.0, 150.0, 75.0, 150.0, shape);
    ASPOSE_ASSERT_EQ(40.0, shape->get_DistanceBottom());
    ASPOSE_ASSERT_EQ(40.0, shape->get_DistanceLeft());
    ASPOSE_ASSERT_EQ(40.0, shape->get_DistanceRight());
    ASPOSE_ASSERT_EQ(40.0, shape->get_DistanceTop());
    ASPOSE_ASSERT_EQ(60.0, shape->get_Rotation());
}

namespace gtest_test
{

TEST_F(ExShape, Coordinates)
{
    s_instance->Coordinates();
}

} // namespace gtest_test

void ExShape::GroupShape()
{
    //ExStart
    //ExFor:ShapeBase.Bounds
    //ExFor:ShapeBase.CoordOrigin
    //ExFor:ShapeBase.CoordSize
    //ExSummary:Shows how to create and populate a group shape.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create a group shape. A group shape can display a collection of child shape nodes.
    // In Microsoft Word, clicking within the group shape's boundary or on one of the group shape's child shapes will
    // select all the other child shapes within this group and allow us to scale and move all the shapes at once.
    auto group = System::MakeObject<Aspose::Words::Drawing::GroupShape>(doc);
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::None, group->get_WrapType());
    
    // Create a 400pt x 400pt group shape and place it at the document's floating shape coordinate origin.
    group->set_Bounds(System::Drawing::RectangleF(0.0f, 0.0f, 400.0f, 400.0f));
    
    // Set the group's internal coordinate plane size to 500 x 500pt. 
    // The top left corner of the group will have an x and y coordinate of (0, 0),
    // and the bottom right corner will have an x and y coordinate of (500, 500).
    group->set_CoordSize(System::Drawing::Size(500, 500));
    
    // Set the coordinates of the top left corner of the group to (-250, -250). 
    // The group's center will now have an x and y coordinate value of (0, 0),
    // and the bottom right corner will be at (250, 250).
    group->set_CoordOrigin(System::Drawing::Point(-250, -250));
    
    // Create a rectangle that will display the boundary of this group shape and add it to the group.
    auto child1 = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Rectangle);
    child1->set_Width(group->get_CoordSize().get_Width());
    child1->set_Height(group->get_CoordSize().get_Height());
    child1->set_Left(group->get_CoordOrigin().get_X());
    child1->set_Top(group->get_CoordOrigin().get_Y());
    group->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(child1);
    
    // Once a shape is a part of a group shape, we can access it as a child node and then modify it.
    (System::ExplicitCast<Aspose::Words::Drawing::Shape>(group->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Stroke()->set_DashStyle(Aspose::Words::Drawing::DashStyle::Dash);
    
    // Create a small red star and insert it into the group.
    // Line up the shape with the group's coordinate origin, which we have moved to the center.
    auto child2 = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Star);
    child2->set_Width(20);
    child2->set_Height(20);
    child2->set_Left(-10);
    child2->set_Top(-10);
    child2->set_FillColor(System::Drawing::Color::get_Red());
    group->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(child2);
    
    // Insert a rectangle, and then insert a slightly smaller rectangle in the same place with an image.
    // Newer shapes that we add to the group overlap older shapes. The light blue rectangle will partially overlap the red star,
    // and then the shape with the image will overlap the light blue rectangle, using it as a frame.
    // We cannot use the "ZOrder" properties of shapes to manipulate their arrangement within a group shape.
    auto child3 = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Rectangle);
    child3->set_Width(250);
    child3->set_Height(250);
    child3->set_Left(-250);
    child3->set_Top(-250);
    child3->set_FillColor(System::Drawing::Color::get_LightBlue());
    group->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(child3);
    
    auto child4 = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Image);
    child4->set_Width(200);
    child4->set_Height(200);
    child4->set_Left(-225);
    child4->set_Top(-225);
    group->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(child4);
    
    (System::ExplicitCast<Aspose::Words::Drawing::Shape>(group->GetChild(Aspose::Words::NodeType::Shape, 3, true)))->get_ImageData()->SetImage(get_ImageDir() + u"Logo.jpg");
    
    // Insert a text box into the group shape. Set the "Left" property so that the text box's right edge
    // touches the right boundary of the group shape. Set the "Top" property so that the text box sits outside
    // the boundary of the group shape, with its top size lined up along the group shape's bottom margin.
    auto child5 = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::TextBox);
    child5->set_Width(200);
    child5->set_Height(50);
    child5->set_Left(group->get_CoordSize().get_Width() + group->get_CoordOrigin().get_X() - 200);
    child5->set_Top(group->get_CoordSize().get_Height() + group->get_CoordOrigin().get_Y());
    group->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(child5);
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->InsertNode(group);
    builder->MoveTo((System::ExplicitCast<Aspose::Words::Drawing::Shape>(group->GetChild(Aspose::Words::NodeType::Shape, 4, true)))->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(System::MakeObject<Aspose::Words::Paragraph>(doc)));
    builder->Write(u"Hello world!");
    
    doc->Save(get_ArtifactsDir() + u"Shape.GroupShape.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.GroupShape.docx");
    group = System::ExplicitCast<Aspose::Words::Drawing::GroupShape>(doc->GetChild(Aspose::Words::NodeType::GroupShape, 0, true));
    
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 0.0f, 400.0f, 400.0f), group->get_Bounds());
    ASPOSE_ASSERT_EQ(System::Drawing::Size(500, 500), group->get_CoordSize());
    ASPOSE_ASSERT_EQ(System::Drawing::Point(-250, -250), group->get_CoordOrigin());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, System::String::Empty, 500.0, 500.0, -250.0, -250.0, System::ExplicitCast<Aspose::Words::Drawing::Shape>(group->GetChild(Aspose::Words::NodeType::Shape, 0, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Star, System::String::Empty, 20.0, 20.0, -10.0, -10.0, System::ExplicitCast<Aspose::Words::Drawing::Shape>(group->GetChild(Aspose::Words::NodeType::Shape, 1, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, System::String::Empty, 250.0, 250.0, -250.0, -250.0, System::ExplicitCast<Aspose::Words::Drawing::Shape>(group->GetChild(Aspose::Words::NodeType::Shape, 2, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Image, System::String::Empty, 200.0, 200.0, -225.0, -225.0, System::ExplicitCast<Aspose::Words::Drawing::Shape>(group->GetChild(Aspose::Words::NodeType::Shape, 3, true)));
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, System::String::Empty, 200.0, 50.0, 250.0, 50.0, System::ExplicitCast<Aspose::Words::Drawing::Shape>(group->GetChild(Aspose::Words::NodeType::Shape, 4, true)));
}

namespace gtest_test
{

TEST_F(ExShape, GroupShape)
{
    s_instance->GroupShape();
}

} // namespace gtest_test

void ExShape::IsTopLevel()
{
    //ExStart
    //ExFor:ShapeBase.IsTopLevel
    //ExSummary:Shows how to tell whether a shape is a part of a group shape.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto shape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Rectangle);
    shape->set_Width(200);
    shape->set_Height(200);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    
    // A shape by default is not part of any group shape, and therefore has the "IsTopLevel" property set to "true".
    ASSERT_TRUE(shape->get_IsTopLevel());
    
    auto group = System::MakeObject<Aspose::Words::Drawing::GroupShape>(doc);
    group->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(shape);
    
    // Once we assimilate a shape into a group shape, the "IsTopLevel" property changes to "false".
    ASSERT_FALSE(shape->get_IsTopLevel());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, IsTopLevel)
{
    s_instance->IsTopLevel();
}

} // namespace gtest_test

void ExShape::LocalToParent()
{
    //ExStart
    //ExFor:ShapeBase.CoordOrigin
    //ExFor:ShapeBase.CoordSize
    //ExFor:ShapeBase.LocalToParent(PointF)
    //ExSummary:Shows how to translate the x and y coordinate location on a shape's coordinate plane to a location on the parent shape's coordinate plane.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Insert a group shape, and place it 100 points below and to the right of
    // the document's x and Y coordinate origin point.
    auto group = System::MakeObject<Aspose::Words::Drawing::GroupShape>(doc);
    group->set_Bounds(System::Drawing::RectangleF(100.0f, 100.0f, 500.0f, 500.0f));
    
    // Use the "LocalToParent" method to determine that (0, 0) on the group's internal x and y coordinates
    // lies on (100, 100) of its parent shape's coordinate system. The group shape's parent is the document itself.
    ASPOSE_ASSERT_EQ(System::Drawing::PointF(100.0f, 100.0f), group->LocalToParent(System::Drawing::PointF(0.0f, 0.0f)));
    
    // By default, a shape's internal coordinate plane has the top left corner at (0, 0),
    // and the bottom right corner at (1000, 1000). Due to its size, our group shape covers an area of 500pt x 500pt
    // in the document's plane. This means that a movement of 1pt on the document's coordinate plane will translate
    // to a movement of 2pts on the group shape's coordinate plane.
    ASPOSE_ASSERT_EQ(System::Drawing::PointF(150.0f, 150.0f), group->LocalToParent(System::Drawing::PointF(100.0f, 100.0f)));
    ASPOSE_ASSERT_EQ(System::Drawing::PointF(200.0f, 200.0f), group->LocalToParent(System::Drawing::PointF(200.0f, 200.0f)));
    ASPOSE_ASSERT_EQ(System::Drawing::PointF(250.0f, 250.0f), group->LocalToParent(System::Drawing::PointF(300.0f, 300.0f)));
    
    // Move the group shape's x and y axis origin from the top left corner to the center.
    // This will offset the group's internal coordinates relative to the document's coordinates even further.
    group->set_CoordOrigin(System::Drawing::Point(-250, -250));
    
    ASPOSE_ASSERT_EQ(System::Drawing::PointF(375.0f, 375.0f), group->LocalToParent(System::Drawing::PointF(300.0f, 300.0f)));
    
    // Changing the scale of the coordinate plane will also affect relative locations.
    group->set_CoordSize(System::Drawing::Size(500, 500));
    
    ASPOSE_ASSERT_EQ(System::Drawing::PointF(650.0f, 650.0f), group->LocalToParent(System::Drawing::PointF(300.0f, 300.0f)));
    
    // If we wish to add a shape to this group while defining its location based on a location in the document,
    // we will need to first confirm a location in the group shape that will match the document's location.
    ASPOSE_ASSERT_EQ(System::Drawing::PointF(700.0f, 700.0f), group->LocalToParent(System::Drawing::PointF(350.0f, 350.0f)));
    
    auto shape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Rectangle);
    shape->set_Width(100);
    shape->set_Height(100);
    shape->set_Left(700);
    shape->set_Top(700);
    
    group->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(shape);
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Drawing::GroupShape>>(group);
    
    doc->Save(get_ArtifactsDir() + u"Shape.LocalToParent.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.LocalToParent.docx");
    group = System::ExplicitCast<Aspose::Words::Drawing::GroupShape>(doc->GetChild(Aspose::Words::NodeType::GroupShape, 0, true));
    
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(100.0f, 100.0f, 500.0f, 500.0f), group->get_Bounds());
    ASPOSE_ASSERT_EQ(System::Drawing::Size(500, 500), group->get_CoordSize());
    ASPOSE_ASSERT_EQ(System::Drawing::Point(-250, -250), group->get_CoordOrigin());
}

namespace gtest_test
{

TEST_F(ExShape, LocalToParent)
{
    s_instance->LocalToParent();
}

} // namespace gtest_test

void ExShape::AnchorLocked(bool anchorLocked)
{
    //ExStart
    //ExFor:ShapeBase.AnchorLocked
    //ExSummary:Shows how to lock or unlock a shape's paragraph anchor.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Hello world!");
    
    builder->Write(u"Our shape will have an anchor attached to this paragraph.");
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 200, 160);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    builder->InsertBreak(Aspose::Words::BreakType::ParagraphBreak);
    
    builder->Writeln(u"Hello again!");
    
    // Set the "AnchorLocked" property to "true" to prevent the shape's anchor
    // from moving when moving the shape in Microsoft Word.
    // Set the "AnchorLocked" property to "false" to allow any movement of the shape
    // to also move its anchor to any other paragraph that the shape ends up close to.
    shape->set_AnchorLocked(anchorLocked);
    
    // If the shape does not have a visible anchor symbol to its left,
    // we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
    doc->Save(get_ArtifactsDir() + u"Shape.AnchorLocked.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.AnchorLocked.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASPOSE_ASSERT_EQ(anchorLocked, shape->get_AnchorLocked());
}

namespace gtest_test
{

using ExShape_AnchorLocked_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExShape::AnchorLocked)>::type;

struct ExShape_AnchorLocked : public ExShape, public Aspose::Words::ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_AnchorLocked_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
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

} // namespace gtest_test

void ExShape::DeleteAllShapes()
{
    //ExStart
    //ExFor:Shape
    //ExSummary:Shows how to delete all shapes from a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert two shapes along with a group shape with another shape inside it.
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 400, 200);
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::Star, 300, 300);
    
    auto group = System::MakeObject<Aspose::Words::Drawing::GroupShape>(doc);
    group->set_Bounds(System::Drawing::RectangleF(100.0f, 50.0f, 200.0f, 100.0f));
    group->set_CoordOrigin(System::Drawing::Point(-1000, -500));
    
    auto subShape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Cube);
    subShape->set_Width(500);
    subShape->set_Height(700);
    subShape->set_Left(0);
    subShape->set_Top(0);
    
    group->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(subShape);
    builder->InsertNode(group);
    
    ASSERT_EQ(3, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::GroupShape, true)->get_Count());
    
    // Remove all Shape nodes from the document.
    System::SharedPtr<Aspose::Words::NodeCollection> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);
    shapes->Clear();
    
    // All shapes are gone, but the group shape is still in the document.
    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::GroupShape, true)->get_Count());
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    
    // Remove all group shapes separately.
    System::SharedPtr<Aspose::Words::NodeCollection> groupShapes = doc->GetChildNodes(Aspose::Words::NodeType::GroupShape, true);
    groupShapes->Clear();
    
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::GroupShape, true)->get_Count());
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, DeleteAllShapes)
{
    s_instance->DeleteAllShapes();
}

} // namespace gtest_test

void ExShape::IsInline()
{
    //ExStart
    //ExFor:ShapeBase.IsInline
    //ExSummary:Shows how to determine whether a shape is inline or floating.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are two wrapping types that shapes may have.
    // 1 -  Inline:
    builder->Write(u"Hello world! ");
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 100, 100);
    shape->set_FillColor(System::Drawing::Color::get_LightBlue());
    builder->Write(u" Hello again.");
    
    // An inline shape sits inside a paragraph among other paragraph elements, such as runs of text.
    // In Microsoft Word, we may click and drag the shape to any paragraph as if it is a character.
    // If the shape is large, it will affect vertical paragraph spacing.
    // We cannot move this shape to a place with no paragraph.
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, shape->get_WrapType());
    ASSERT_TRUE(shape->get_IsInline());
    
    // 2 -  Floating:
    shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 200, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 200, 100, 100, Aspose::Words::Drawing::WrapType::None);
    shape->set_FillColor(System::Drawing::Color::get_Orange());
    
    // A floating shape belongs to the paragraph that we insert it into,
    // which we can determine by an anchor symbol that appears when we click the shape.
    // If the shape does not have a visible anchor symbol to its left,
    // we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
    // In Microsoft Word, we may left click and drag this shape freely to any location.
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::None, shape->get_WrapType());
    ASSERT_FALSE(shape->get_IsInline());
    
    doc->Save(get_ArtifactsDir() + u"Shape.IsInline.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.IsInline.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, u"Rectangle 100002", 100, 100, 0, 0, shape);
    ASSERT_EQ(System::Drawing::Color::get_LightBlue().ToArgb(), shape->get_FillColor().ToArgb());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, shape->get_WrapType());
    ASSERT_TRUE(shape->get_IsInline());
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, u"Rectangle 100004", 100, 100, 200, 200, shape);
    ASSERT_EQ(System::Drawing::Color::get_Orange().ToArgb(), shape->get_FillColor().ToArgb());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::None, shape->get_WrapType());
    ASSERT_FALSE(shape->get_IsInline());
}

namespace gtest_test
{

TEST_F(ExShape, IsInline)
{
    s_instance->IsInline();
}

} // namespace gtest_test

void ExShape::Bounds()
{
    //ExStart
    //ExFor:ShapeBase.Bounds
    //ExFor:ShapeBase.BoundsInPoints
    //ExSummary:Shows how to verify shape containing block boundaries.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Line, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 50, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 50, 100, 100, Aspose::Words::Drawing::WrapType::None);
    shape->set_StrokeColor(System::Drawing::Color::get_Orange());
    
    // Even though the line itself takes up little space on the document page,
    // it occupies a rectangular containing block, the size of which we can determine using the "Bounds" properties.
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(50.0f, 50.0f, 100.0f, 100.0f), shape->get_Bounds());
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(50.0f, 50.0f, 100.0f, 100.0f), shape->get_BoundsInPoints());
    
    // Create a group shape, and then set the size of its containing block using the "Bounds" property.
    auto group = System::MakeObject<Aspose::Words::Drawing::GroupShape>(doc);
    group->set_Bounds(System::Drawing::RectangleF(0.0f, 100.0f, 250.0f, 250.0f));
    
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 100.0f, 250.0f, 250.0f), group->get_BoundsInPoints());
    
    // Create a rectangle, verify the size of its bounding block, and then add it to the group shape.
    shape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Rectangle);
    shape->set_Width(100);
    shape->set_Height(100);
    shape->set_Left(700);
    shape->set_Top(700);
    
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(700.0f, 700.0f, 100.0f, 100.0f), shape->get_BoundsInPoints());
    
    group->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(shape);
    
    // The group shape's coordinate plane has its origin on the top left-hand side corner of its containing block,
    // and the x and y coordinates of (1000, 1000) on the bottom right-hand side corner.
    // Our group shape is 250x250pt in size, so every 4pt on the group shape's coordinate plane
    // translates to 1pt in the document body's coordinate plane.
    // Every shape that we insert will also shrink in size by a factor of 4.
    // The change in the shape's "BoundsInPoints" property will reflect this.
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(175.0f, 275.0f, 25.0f, 25.0f), shape->get_BoundsInPoints());
    
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Drawing::GroupShape>>(group);
    
    // Insert a shape and place it outside of the bounds of the group shape's containing block.
    shape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Rectangle);
    shape->set_Width(100);
    shape->set_Height(100);
    shape->set_Left(1000);
    shape->set_Top(1000);
    
    group->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(shape);
    
    // The group shape's footprint in the document body has increased, but the containing block remains the same.
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 100.0f, 250.0f, 250.0f), group->get_BoundsInPoints());
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(250.0f, 350.0f, 25.0f, 25.0f), shape->get_BoundsInPoints());
    
    doc->Save(get_ArtifactsDir() + u"Shape.Bounds.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.Bounds.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Line, u"Line 100002", 100, 100, 50, 50, shape);
    ASSERT_EQ(System::Drawing::Color::get_Orange().ToArgb(), shape->get_StrokeColor().ToArgb());
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(50.0f, 50.0f, 100.0f, 100.0f), shape->get_BoundsInPoints());
    
    group = System::ExplicitCast<Aspose::Words::Drawing::GroupShape>(doc->GetChild(Aspose::Words::NodeType::GroupShape, 0, true));
    
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 100.0f, 250.0f, 250.0f), group->get_Bounds());
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 100.0f, 250.0f, 250.0f), group->get_BoundsInPoints());
    ASPOSE_ASSERT_EQ(System::Drawing::Size(1000, 1000), group->get_CoordSize());
    ASPOSE_ASSERT_EQ(System::Drawing::Point(0, 0), group->get_CoordOrigin());
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, System::String::Empty, 100, 100, 700, 700, shape);
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(175.0f, 275.0f, 25.0f, 25.0f), shape->get_BoundsInPoints());
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, System::String::Empty, 100, 100, 1000, 1000, shape);
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(250.0f, 350.0f, 25.0f, 25.0f), shape->get_BoundsInPoints());
}

namespace gtest_test
{

TEST_F(ExShape, Bounds)
{
    s_instance->Bounds();
}

} // namespace gtest_test

void ExShape::FlipShapeOrientation()
{
    //ExStart
    //ExFor:ShapeBase.FlipOrientation
    //ExFor:FlipOrientation
    //ExSummary:Shows how to flip a shape on an axis.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert an image shape and leave its orientation in its default state.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 100, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 100, 100, 100, Aspose::Words::Drawing::WrapType::None);
    shape->get_ImageData()->SetImage(get_ImageDir() + u"Logo.jpg");
    
    ASSERT_EQ(Aspose::Words::Drawing::FlipOrientation::None, shape->get_FlipOrientation());
    
    shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 250, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 100, 100, 100, Aspose::Words::Drawing::WrapType::None);
    shape->get_ImageData()->SetImage(get_ImageDir() + u"Logo.jpg");
    
    // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the second shape on the y-axis,
    // making it into a horizontal mirror image of the first shape.
    shape->set_FlipOrientation(Aspose::Words::Drawing::FlipOrientation::Horizontal);
    
    shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 100, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 250, 100, 100, Aspose::Words::Drawing::WrapType::None);
    shape->get_ImageData()->SetImage(get_ImageDir() + u"Logo.jpg");
    
    // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the third shape on the x-axis,
    // making it into a vertical mirror image of the first shape.
    shape->set_FlipOrientation(Aspose::Words::Drawing::FlipOrientation::Vertical);
    
    shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 250, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 250, 100, 100, Aspose::Words::Drawing::WrapType::None);
    shape->get_ImageData()->SetImage(get_ImageDir() + u"Logo.jpg");
    
    // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the fourth shape on both the x and y axes,
    // making it into a horizontal and vertical mirror image of the first shape.
    shape->set_FlipOrientation(Aspose::Words::Drawing::FlipOrientation::Both);
    
    doc->Save(get_ArtifactsDir() + u"Shape.FlipShapeOrientation.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.FlipShapeOrientation.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, u"Rectangle 100002", 100, 100, 100, 100, shape);
    ASSERT_EQ(Aspose::Words::Drawing::FlipOrientation::None, shape->get_FlipOrientation());
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, u"Rectangle 100004", 100, 100, 100, 250, shape);
    ASSERT_EQ(Aspose::Words::Drawing::FlipOrientation::Horizontal, shape->get_FlipOrientation());
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, u"Rectangle 100006", 100, 100, 250, 100, shape);
    ASSERT_EQ(Aspose::Words::Drawing::FlipOrientation::Vertical, shape->get_FlipOrientation());
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 3, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, u"Rectangle 100008", 100, 100, 250, 250, shape);
    ASSERT_EQ(Aspose::Words::Drawing::FlipOrientation::Both, shape->get_FlipOrientation());
}

namespace gtest_test
{

TEST_F(ExShape, FlipShapeOrientation)
{
    s_instance->FlipShapeOrientation();
}

} // namespace gtest_test

void ExShape::Fill()
{
    //ExStart
    //ExFor:ShapeBase.Fill
    //ExFor:Shape.FillColor
    //ExFor:Shape.StrokeColor
    //ExFor:Fill
    //ExFor:Fill.Opacity
    //ExSummary:Shows how to fill a shape with a solid color.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Write some text, and then cover it with a floating shape.
    builder->get_Font()->set_Size(32);
    builder->Writeln(u"Hello world!");
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::CloudCallout, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 25, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 25, 250, 150, Aspose::Words::Drawing::WrapType::None);
    
    // Use the "StrokeColor" property to set the color of the outline of the shape.
    shape->set_StrokeColor(System::Drawing::Color::get_CadetBlue());
    
    // Use the "FillColor" property to set the color of the inside area of the shape.
    shape->set_FillColor(System::Drawing::Color::get_LightBlue());
    
    // The "Opacity" property determines how transparent the color is on a 0-1 scale,
    // with 1 being fully opaque, and 0 being invisible.
    // The shape fill by default is fully opaque, so we cannot see the text that this shape is on top of.
    ASPOSE_ASSERT_EQ(1.0, shape->get_Fill()->get_Opacity());
    
    // Set the shape fill color's opacity to a lower value so that we can see the text underneath it.
    shape->get_Fill()->set_Opacity(0.3);
    
    doc->Save(get_ArtifactsDir() + u"Shape.Fill.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.Fill.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::CloudCallout, u"CloudCallout 100002", 250.0, 150.0, 25.0, 25.0, shape);
    System::Drawing::Color colorWithOpacity = System::Drawing::Color::FromArgb(System::Convert::ToInt32(255 * shape->get_Fill()->get_Opacity()), System::Drawing::Color::get_LightBlue().get_R(), System::Drawing::Color::get_LightBlue().get_G(), System::Drawing::Color::get_LightBlue().get_B());
    ASSERT_EQ(colorWithOpacity.ToArgb(), shape->get_FillColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_CadetBlue().ToArgb(), shape->get_StrokeColor().ToArgb());
    ASSERT_NEAR(0.3, shape->get_Fill()->get_Opacity(), 0.01);
}

namespace gtest_test
{

TEST_F(ExShape, Fill)
{
    s_instance->Fill();
}

} // namespace gtest_test

void ExShape::TextureFill()
{
    //ExStart
    //ExFor:Fill.PresetTexture
    //ExFor:Fill.TextureAlignment
    //ExFor:TextureAlignment
    //ExSummary:Shows how to fill and tiling the texture inside the shape.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 80, 80);
    
    // Apply texture alignment to the shape fill.
    shape->get_Fill()->PresetTextured(Aspose::Words::Drawing::PresetTexture::Canvas);
    shape->get_Fill()->set_TextureAlignment(Aspose::Words::Drawing::TextureAlignment::TopRight);
    
    // Use the compliance option to define the shape using DML if you want to get "TextureAlignment"
    // property after the document saves.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    saveOptions->set_Compliance(Aspose::Words::Saving::OoxmlCompliance::Iso29500_2008_Strict);
    
    doc->Save(get_ArtifactsDir() + u"Shape.TextureFill.docx", saveOptions);
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.TextureFill.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(Aspose::Words::Drawing::TextureAlignment::TopRight, shape->get_Fill()->get_TextureAlignment());
    ASSERT_EQ(Aspose::Words::Drawing::PresetTexture::Canvas, shape->get_Fill()->get_PresetTexture());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, TextureFill)
{
    s_instance->TextureFill();
}

} // namespace gtest_test

void ExShape::GradientFill()
{
    //ExStart
    //ExFor:Fill.OneColorGradient(Color, GradientStyle, GradientVariant, Double)
    //ExFor:Fill.OneColorGradient(GradientStyle, GradientVariant, Double)
    //ExFor:Fill.TwoColorGradient(Color, Color, GradientStyle, GradientVariant)
    //ExFor:Fill.TwoColorGradient(GradientStyle, GradientVariant)
    //ExFor:Fill.BackColor
    //ExFor:Fill.GradientStyle
    //ExFor:Fill.GradientVariant
    //ExFor:Fill.GradientAngle
    //ExFor:GradientStyle
    //ExFor:GradientVariant
    //ExSummary:Shows how to fill a shape with a gradients.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 80, 80);
    // Apply One-color gradient fill to the shape with ForeColor of gradient fill.
    shape->get_Fill()->OneColorGradient(System::Drawing::Color::get_Red(), Aspose::Words::Drawing::GradientStyle::Horizontal, Aspose::Words::Drawing::GradientVariant::Variant2, 0.1);
    
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), shape->get_Fill()->get_ForeColor().ToArgb());
    ASSERT_EQ(Aspose::Words::Drawing::GradientStyle::Horizontal, shape->get_Fill()->get_GradientStyle());
    ASSERT_EQ(Aspose::Words::Drawing::GradientVariant::Variant2, shape->get_Fill()->get_GradientVariant());
    ASPOSE_ASSERT_EQ(270, shape->get_Fill()->get_GradientAngle());
    
    shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 80, 80);
    // Apply Two-color gradient fill to the shape.
    shape->get_Fill()->TwoColorGradient(Aspose::Words::Drawing::GradientStyle::FromCorner, Aspose::Words::Drawing::GradientVariant::Variant4);
    // Change BackColor of gradient fill.
    shape->get_Fill()->set_BackColor(System::Drawing::Color::get_Yellow());
    // Note that changes "GradientAngle" for "GradientStyle.FromCorner/GradientStyle.FromCenter"
    // gradient fill don't get any effect, it will work only for linear gradient.
    shape->get_Fill()->set_GradientAngle(15);
    
    ASSERT_EQ(System::Drawing::Color::get_Yellow().ToArgb(), shape->get_Fill()->get_BackColor().ToArgb());
    ASSERT_EQ(Aspose::Words::Drawing::GradientStyle::FromCorner, shape->get_Fill()->get_GradientStyle());
    ASSERT_EQ(Aspose::Words::Drawing::GradientVariant::Variant4, shape->get_Fill()->get_GradientVariant());
    ASPOSE_ASSERT_EQ(0, shape->get_Fill()->get_GradientAngle());
    
    // Use the compliance option to define the shape using DML if you want to get "GradientStyle",
    // "GradientVariant" and "GradientAngle" properties after the document saves.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    saveOptions->set_Compliance(Aspose::Words::Saving::OoxmlCompliance::Iso29500_2008_Strict);
    
    doc->Save(get_ArtifactsDir() + u"Shape.GradientFill.docx", saveOptions);
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.GradientFill.docx");
    auto firstShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), firstShape->get_Fill()->get_ForeColor().ToArgb());
    ASSERT_EQ(Aspose::Words::Drawing::GradientStyle::Horizontal, firstShape->get_Fill()->get_GradientStyle());
    ASSERT_EQ(Aspose::Words::Drawing::GradientVariant::Variant2, firstShape->get_Fill()->get_GradientVariant());
    ASPOSE_ASSERT_EQ(270, firstShape->get_Fill()->get_GradientAngle());
    
    auto secondShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    ASSERT_EQ(System::Drawing::Color::get_Yellow().ToArgb(), secondShape->get_Fill()->get_BackColor().ToArgb());
    ASSERT_EQ(Aspose::Words::Drawing::GradientStyle::FromCorner, secondShape->get_Fill()->get_GradientStyle());
    ASSERT_EQ(Aspose::Words::Drawing::GradientVariant::Variant4, secondShape->get_Fill()->get_GradientVariant());
    ASPOSE_ASSERT_EQ(0, secondShape->get_Fill()->get_GradientAngle());
}

namespace gtest_test
{

TEST_F(ExShape, GradientFill)
{
    s_instance->GradientFill();
}

} // namespace gtest_test

void ExShape::GradientStops()
{
    //ExStart
    //ExFor:Fill.GradientStops
    //ExFor:GradientStopCollection
    //ExFor:GradientStopCollection.Insert(Int32, GradientStop)
    //ExFor:GradientStopCollection.Add(GradientStop)
    //ExFor:GradientStopCollection.RemoveAt(Int32)
    //ExFor:GradientStopCollection.Remove(GradientStop)
    //ExFor:GradientStopCollection.Item(Int32)
    //ExFor:GradientStopCollection.Count
    //ExFor:GradientStop
    //ExFor:GradientStop.#ctor(Color, Double)
    //ExFor:GradientStop.#ctor(Color, Double, Double)
    //ExFor:GradientStop.BaseColor
    //ExFor:GradientStop.Color
    //ExFor:GradientStop.Position
    //ExFor:GradientStop.Transparency
    //ExFor:GradientStop.Remove
    //ExSummary:Shows how to add gradient stops to the gradient fill.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 80, 80);
    shape->get_Fill()->TwoColorGradient(System::Drawing::Color::get_Green(), System::Drawing::Color::get_Red(), Aspose::Words::Drawing::GradientStyle::Horizontal, Aspose::Words::Drawing::GradientVariant::Variant2);
    
    // Get gradient stops collection.
    System::SharedPtr<Aspose::Words::Drawing::GradientStopCollection> gradientStops = shape->get_Fill()->get_GradientStops();
    
    // Change first gradient stop.
    gradientStops->idx_get(0)->set_Color(System::Drawing::Color::get_Aqua());
    gradientStops->idx_get(0)->set_Position(0.1);
    gradientStops->idx_get(0)->set_Transparency(0.25);
    
    // Add new gradient stop to the end of collection.
    auto gradientStop = System::MakeObject<Aspose::Words::Drawing::GradientStop>(System::Drawing::Color::get_Brown(), 0.5);
    gradientStops->Add(gradientStop);
    
    // Remove gradient stop at index 1.
    gradientStops->RemoveAt(1);
    // And insert new gradient stop at the same index 1.
    gradientStops->Insert(1, System::MakeObject<Aspose::Words::Drawing::GradientStop>(System::Drawing::Color::get_Chocolate(), 0.75, 0.3));
    
    // Remove last gradient stop in the collection.
    gradientStop = gradientStops->idx_get(2);
    gradientStops->Remove(gradientStop);
    
    ASSERT_EQ(2, gradientStops->get_Count());
    
    ASPOSE_ASSERT_EQ(System::Drawing::Color::FromArgb(255, 0, 255, 255), gradientStops->idx_get(0)->get_BaseColor());
    ASSERT_EQ(System::Drawing::Color::get_Aqua().ToArgb(), gradientStops->idx_get(0)->get_Color().ToArgb());
    ASSERT_NEAR(0.1, gradientStops->idx_get(0)->get_Position(), 0.01);
    ASSERT_NEAR(0.25, gradientStops->idx_get(0)->get_Transparency(), 0.01);
    
    ASSERT_EQ(System::Drawing::Color::get_Chocolate().ToArgb(), gradientStops->idx_get(1)->get_Color().ToArgb());
    ASSERT_NEAR(0.75, gradientStops->idx_get(1)->get_Position(), 0.01);
    ASSERT_NEAR(0.3, gradientStops->idx_get(1)->get_Transparency(), 0.01);
    
    // Use the compliance option to define the shape using DML
    // if you want to get "GradientStops" property after the document saves.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    saveOptions->set_Compliance(Aspose::Words::Saving::OoxmlCompliance::Iso29500_2008_Strict);
    
    doc->Save(get_ArtifactsDir() + u"Shape.GradientStops.docx", saveOptions);
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.GradientStops.docx");
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    gradientStops = shape->get_Fill()->get_GradientStops();
    
    ASSERT_EQ(2, gradientStops->get_Count());
    
    ASSERT_EQ(System::Drawing::Color::get_Aqua().ToArgb(), gradientStops->idx_get(0)->get_Color().ToArgb());
    ASSERT_NEAR(0.1, gradientStops->idx_get(0)->get_Position(), 0.01);
    ASSERT_NEAR(0.25, gradientStops->idx_get(0)->get_Transparency(), 0.01);
    
    ASSERT_EQ(System::Drawing::Color::get_Chocolate().ToArgb(), gradientStops->idx_get(1)->get_Color().ToArgb());
    ASSERT_NEAR(0.75, gradientStops->idx_get(1)->get_Position(), 0.01);
    ASSERT_NEAR(0.3, gradientStops->idx_get(1)->get_Transparency(), 0.01);
}

namespace gtest_test
{

TEST_F(ExShape, GradientStops)
{
    s_instance->GradientStops();
}

} // namespace gtest_test

void ExShape::FillPattern()
{
    //ExStart
    //ExFor:PatternType
    //ExFor:Fill.Pattern
    //ExFor:Fill.Patterned(PatternType)
    //ExFor:Fill.Patterned(PatternType, Color, Color)
    //ExSummary:Shows how to set pattern for a shape.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Shape stroke pattern border.docx");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    System::SharedPtr<Aspose::Words::Drawing::Fill> fill = shape->get_Fill();
    
    std::cout << System::String::Format(u"Pattern value is: {0}", fill->get_Pattern()) << std::endl;
    
    // There are several ways specified fill to a pattern.
    // 1 -  Apply pattern to the shape fill:
    fill->Patterned(Aspose::Words::Drawing::PatternType::DiagonalBrick);
    
    // 2 -  Apply pattern with foreground and background colors to the shape fill:
    fill->Patterned(Aspose::Words::Drawing::PatternType::DiagonalBrick, System::Drawing::Color::get_Aqua(), System::Drawing::Color::get_Bisque());
    
    doc->Save(get_ArtifactsDir() + u"Shape.FillPattern.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, FillPattern)
{
    s_instance->FillPattern();
}

} // namespace gtest_test

void ExShape::FillThemeColor()
{
    //ExStart
    //ExFor:Fill.ForeThemeColor
    //ExFor:Fill.BackThemeColor
    //ExFor:Fill.BackTintAndShade
    //ExSummary:Shows how to set theme color for foreground/background shape color.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::RoundRectangle, 80, 80);
    
    System::SharedPtr<Aspose::Words::Drawing::Fill> fill = shape->get_Fill();
    fill->set_ForeThemeColor(Aspose::Words::Themes::ThemeColor::Dark1);
    fill->set_BackThemeColor(Aspose::Words::Themes::ThemeColor::Background2);
    
    // Note: do not use "BackThemeColor" and "BackTintAndShade" for font fill.
    if (fill->get_BackTintAndShade() == 0)
    {
        fill->set_BackTintAndShade(0.2);
    }
    
    doc->Save(get_ArtifactsDir() + u"Shape.FillThemeColor.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, FillThemeColor)
{
    s_instance->FillThemeColor();
}

} // namespace gtest_test

void ExShape::FillTintAndShade()
{
    //ExStart
    //ExFor:Fill.ForeTintAndShade
    //ExSummary:Shows how to manage lightening and darkening foreground font color.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    
    System::SharedPtr<Aspose::Words::Drawing::Fill> textFill = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Fill();
    textFill->set_ForeThemeColor(Aspose::Words::Themes::ThemeColor::Accent1);
    if (textFill->get_ForeTintAndShade() == 0)
    {
        textFill->set_ForeTintAndShade(0.5);
    }
    
    doc->Save(get_ArtifactsDir() + u"Shape.FillTintAndShade.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, FillTintAndShade)
{
    s_instance->FillTintAndShade();
}

} // namespace gtest_test

void ExShape::Title()
{
    //ExStart
    //ExFor:ShapeBase.Title
    //ExSummary:Shows how to set the title of a shape.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create a shape, give it a title, and then add it to the document.
    auto shape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Cube);
    shape->set_Width(200);
    shape->set_Height(200);
    shape->set_Title(u"My cube");
    
    builder->InsertNode(shape);
    
    // When we save a document with a shape that has a title,
    // Aspose.Words will store that title in the shape's Alt Text.
    doc->Save(get_ArtifactsDir() + u"Shape.Title.docx");
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.Title.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(System::String::Empty, shape->get_Title());
    ASSERT_EQ(u"Title: My cube", shape->get_AlternativeText());
    //ExEnd
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Cube, System::String::Empty, 200.0, 200.0, 0.0, 0.0, shape);
}

namespace gtest_test
{

TEST_F(ExShape, Title)
{
    s_instance->Title();
}

} // namespace gtest_test

void ExShape::ReplaceTextboxesWithImages()
{
    //ExStart
    //ExFor:WrapSide
    //ExFor:ShapeBase.WrapSide
    //ExFor:NodeCollection
    //ExFor:CompositeNode.InsertAfter``1(``0,Node)
    //ExFor:NodeCollection.ToArray
    //ExSummary:Shows how to replace all textbox shapes with image shapes.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Textboxes in drawing canvas.docx");
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Drawing::Shape>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToArray();
    
    ASSERT_EQ(3, shapes->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Drawing::Shape>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Drawing::Shape> s)>>([](System::SharedPtr<Aspose::Words::Drawing::Shape> s) -> bool
    {
        return s->get_ShapeType() == Aspose::Words::Drawing::ShapeType::TextBox;
    }))));
    ASSERT_EQ(1, shapes->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Drawing::Shape>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Drawing::Shape> s)>>([](System::SharedPtr<Aspose::Words::Drawing::Shape> s) -> bool
    {
        return s->get_ShapeType() == Aspose::Words::Drawing::ShapeType::Image;
    }))));
    
    for (System::SharedPtr<Aspose::Words::Drawing::Shape> shape : shapes)
    {
        if (shape->get_ShapeType() == Aspose::Words::Drawing::ShapeType::TextBox)
        {
            auto replacementShape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Image);
            replacementShape->get_ImageData()->SetImage(get_ImageDir() + u"Logo.jpg");
            replacementShape->set_Left(shape->get_Left());
            replacementShape->set_Top(shape->get_Top());
            replacementShape->set_Width(shape->get_Width());
            replacementShape->set_Height(shape->get_Height());
            replacementShape->set_RelativeHorizontalPosition(shape->get_RelativeHorizontalPosition());
            replacementShape->set_RelativeVerticalPosition(shape->get_RelativeVerticalPosition());
            replacementShape->set_HorizontalAlignment(shape->get_HorizontalAlignment());
            replacementShape->set_VerticalAlignment(shape->get_VerticalAlignment());
            replacementShape->set_WrapType(shape->get_WrapType());
            replacementShape->set_WrapSide(shape->get_WrapSide());
            
            shape->get_ParentNode()->InsertAfter<System::SharedPtr<Aspose::Words::Drawing::Shape>>(replacementShape, shape);
            shape->Remove();
        }
    }
    
    
    shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToArray();
    
    ASSERT_EQ(0, shapes->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Drawing::Shape>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Drawing::Shape> s)>>([](System::SharedPtr<Aspose::Words::Drawing::Shape> s) -> bool
    {
        return s->get_ShapeType() == Aspose::Words::Drawing::ShapeType::TextBox;
    }))));
    ASSERT_EQ(4, shapes->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Drawing::Shape>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Drawing::Shape> s)>>([](System::SharedPtr<Aspose::Words::Drawing::Shape> s) -> bool
    {
        return s->get_ShapeType() == Aspose::Words::Drawing::ShapeType::Image;
    }))));
    
    doc->Save(get_ArtifactsDir() + u"Shape.ReplaceTextboxesWithImages.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.ReplaceTextboxesWithImages.docx");
    auto outShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapSide::Both, outShape->get_WrapSide());
}

namespace gtest_test
{

TEST_F(ExShape, ReplaceTextboxesWithImages)
{
    s_instance->ReplaceTextboxesWithImages();
}

} // namespace gtest_test

void ExShape::CreateTextBox()
{
    //ExStart
    //ExFor:Shape.#ctor(DocumentBase, ShapeType)
    //ExFor:Story.FirstParagraph
    //ExFor:Shape.FirstParagraph
    //ExFor:ShapeBase.WrapType
    //ExSummary:Shows how to create and format a text box.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create a floating text box.
    auto textBox = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::TextBox);
    textBox->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    textBox->set_Height(50);
    textBox->set_Width(200);
    
    // Set the horizontal, and vertical alignment of the text inside the shape.
    textBox->set_HorizontalAlignment(Aspose::Words::Drawing::HorizontalAlignment::Center);
    textBox->set_VerticalAlignment(Aspose::Words::Drawing::VerticalAlignment::Top);
    
    // Add a paragraph to the text box and add a run of text that the text box will display.
    textBox->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(System::MakeObject<Aspose::Words::Paragraph>(doc));
    System::SharedPtr<Aspose::Words::Paragraph> para = textBox->get_FirstParagraph();
    para->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    auto run = System::MakeObject<Aspose::Words::Run>(doc);
    run->set_Text(u"Hello world!");
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(textBox);
    
    doc->Save(get_ArtifactsDir() + u"Shape.CreateTextBox.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.CreateTextBox.docx");
    textBox = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, System::String::Empty, 200.0, 50.0, 0.0, 0.0, textBox);
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::None, textBox->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::HorizontalAlignment::Center, textBox->get_HorizontalAlignment());
    ASSERT_EQ(Aspose::Words::Drawing::VerticalAlignment::Top, textBox->get_VerticalAlignment());
    ASSERT_EQ(u"Hello world!", textBox->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExShape, CreateTextBox)
{
    s_instance->CreateTextBox();
}

} // namespace gtest_test

void ExShape::ZOrder()
{
    //ExStart
    //ExFor:ShapeBase.ZOrder
    //ExSummary:Shows how to manipulate the order of shapes.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert three different colored rectangles that partially overlap each other.
    // When we insert a shape that overlaps another shape, Aspose.Words places the newer shape on top of the old one.
    // The light green rectangle will overlap the light blue rectangle and partially obscure it,
    // and the light blue rectangle will obscure the orange rectangle.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 100, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 100, 200, 200, Aspose::Words::Drawing::WrapType::None);
    shape->set_FillColor(System::Drawing::Color::get_Orange());
    
    shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 150, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 150, 200, 200, Aspose::Words::Drawing::WrapType::None);
    shape->set_FillColor(System::Drawing::Color::get_LightBlue());
    
    shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 200, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 200, 200, 200, Aspose::Words::Drawing::WrapType::None);
    shape->set_FillColor(System::Drawing::Color::get_LightGreen());
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Drawing::Shape>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToArray();
    
    // The "ZOrder" property of a shape determines its stacking priority among other overlapping shapes.
    // If two overlapping shapes have different "ZOrder" values,
    // Microsoft Word will place the shape with a higher value over the shape with the lower value. 
    // Set the "ZOrder" values of our shapes to place the first orange rectangle over the second light blue one
    // and the second light blue rectangle over the third light green rectangle.
    // This will reverse their original stacking order.
    shapes[0]->set_ZOrder(3);
    shapes[1]->set_ZOrder(2);
    shapes[2]->set_ZOrder(1);
    
    doc->Save(get_ArtifactsDir() + u"Shape.ZOrder.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, ZOrder)
{
    s_instance->ZOrder();
}

} // namespace gtest_test

void ExShape::GetActiveXControlProperties()
{
    //ExStart
    //ExFor:OleControl
    //ExFor:OleControl.IsForms2OleControl
    //ExFor:OleControl.Name
    //ExFor:OleFormat.OleControl
    //ExFor:Forms2OleControl
    //ExFor:Forms2OleControl.Caption
    //ExFor:Forms2OleControl.Value
    //ExFor:Forms2OleControl.Enabled
    //ExFor:Forms2OleControl.Type
    //ExFor:Forms2OleControl.ChildNodes
    //ExFor:Forms2OleControl.GroupName
    //ExSummary:Shows how to verify the properties of an ActiveX control.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"ActiveX controls.docx");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    System::SharedPtr<Aspose::Words::Drawing::Ole::OleControl> oleControl = shape->get_OleFormat()->get_OleControl();
    
    ASSERT_EQ(u"CheckBox1", oleControl->get_Name());
    
    if (oleControl->get_IsForms2OleControl())
    {
        auto checkBox = System::ExplicitCast<Aspose::Words::Drawing::Ole::Forms2OleControl>(oleControl);
        ASSERT_EQ(u"First", checkBox->get_Caption());
        ASSERT_EQ(u"0", checkBox->get_Value());
        ASPOSE_ASSERT_EQ(true, checkBox->get_Enabled());
        ASSERT_EQ(Aspose::Words::Drawing::Ole::Forms2OleControlType::CheckBox, checkBox->get_Type());
        ASPOSE_ASSERT_EQ(nullptr, checkBox->get_ChildNodes());
        ASSERT_EQ(System::String::Empty, checkBox->get_GroupName());
        
        // Note, that you can't set GroupName for a Frame.
        checkBox->set_GroupName(u"Aspose group name");
    }
    //ExEnd
    
    doc->Save(get_ArtifactsDir() + u"Shape.GetActiveXControlProperties.docx");
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.GetActiveXControlProperties.docx");
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    auto forms2OleControl = System::ExplicitCast<Aspose::Words::Drawing::Ole::Forms2OleControl>(shape->get_OleFormat()->get_OleControl());
    
    ASSERT_EQ(u"Aspose group name", forms2OleControl->get_GroupName());
}

namespace gtest_test
{

TEST_F(ExShape, GetActiveXControlProperties)
{
    s_instance->GetActiveXControlProperties();
}

} // namespace gtest_test

void ExShape::GetOleObjectRawData()
{
    //ExStart
    //ExFor:OleFormat.GetRawData
    //ExSummary:Shows how to access the raw data of an embedded OLE object.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"OLE objects.docx");
    
    for (auto&& shape : System::IterateOver<Aspose::Words::Drawing::Shape>(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)))
    {
        System::SharedPtr<Aspose::Words::Drawing::OleFormat> oleFormat = shape->get_OleFormat();
        if (oleFormat != nullptr)
        {
            std::cout << System::String::Format(u"This is {0} object", (oleFormat->get_IsLink() ? System::String(u"a linked") : System::String(u"an embedded"))) << std::endl;
            System::ArrayPtr<uint8_t> oleRawData = oleFormat->GetRawData();
            
            ASSERT_EQ(24576, oleRawData->get_Length());
        }
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, GetOleObjectRawData)
{
    s_instance->GetOleObjectRawData();
}

} // namespace gtest_test

void ExShape::LinkedChartSourceFullName()
{
    //ExStart
    //ExFor:Chart.SourceFullName
    //ExSummary:Shows how to get/set the full name of the external xls/xlsx document if the chart is linked.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Shape with linked chart.docx");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    System::String sourceFullName = shape->get_Chart()->get_SourceFullName();
    ASSERT_TRUE(sourceFullName.Contains(u"Examples\\Data\\Spreadsheet.xlsx"));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, LinkedChartSourceFullName)
{
    s_instance->LinkedChartSourceFullName();
}

} // namespace gtest_test

void ExShape::OleControl()
{
    //ExStart
    //ExFor:OleFormat
    //ExFor:OleFormat.AutoUpdate
    //ExFor:OleFormat.IsLocked
    //ExFor:OleFormat.ProgId
    //ExFor:OleFormat.Save(Stream)
    //ExFor:OleFormat.Save(String)
    //ExFor:OleFormat.SuggestedExtension
    //ExSummary:Shows how to extract embedded OLE objects into files.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"OLE spreadsheet.docm");
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    // The OLE object in the first shape is a Microsoft Excel spreadsheet.
    System::SharedPtr<Aspose::Words::Drawing::OleFormat> oleFormat = shape->get_OleFormat();
    
    ASSERT_EQ(u"Excel.Sheet.12", oleFormat->get_ProgId());
    
    // Our object is neither auto updating nor locked from updates.
    ASSERT_FALSE(oleFormat->get_AutoUpdate());
    ASPOSE_ASSERT_EQ(false, oleFormat->get_IsLocked());
    
    // If we plan on saving the OLE object to a file in the local file system,
    // we can use the "SuggestedExtension" property to determine which file extension to apply to the file.
    ASSERT_EQ(u".xlsx", oleFormat->get_SuggestedExtension());
    
    // Below are two ways of saving an OLE object to a file in the local file system.
    // 1 -  Save it via a stream:
    {
        auto fs = System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + u"OLE spreadsheet extracted via stream" + oleFormat->get_SuggestedExtension(), System::IO::FileMode::Create);
        oleFormat->Save(fs);
    }
    
    // 2 -  Save it directly to a filename:
    oleFormat->Save(get_ArtifactsDir() + u"OLE spreadsheet saved directly" + oleFormat->get_SuggestedExtension());
    //ExEnd
    
    ASSERT_TRUE(System::MakeObject<System::IO::FileInfo>(get_ArtifactsDir() + u"OLE spreadsheet extracted via stream.xlsx")->get_Length() < 8400);
    ASSERT_TRUE(System::MakeObject<System::IO::FileInfo>(get_ArtifactsDir() + u"OLE spreadsheet saved directly.xlsx")->get_Length() < 8400);
}

namespace gtest_test
{

TEST_F(ExShape, OleControl)
{
    s_instance->OleControl();
}

} // namespace gtest_test

void ExShape::OleLinks()
{
    //ExStart
    //ExFor:OleFormat.IconCaption
    //ExFor:OleFormat.GetOleEntry(String)
    //ExFor:OleFormat.IsLink
    //ExFor:OleFormat.OleIcon
    //ExFor:OleFormat.SourceFullName
    //ExFor:OleFormat.SourceItem
    //ExSummary:Shows how to insert linked and unlinked OLE objects.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Embed a Microsoft Visio drawing into the document as an OLE object.
    builder->InsertOleObject(get_ImageDir() + u"Microsoft Visio drawing.vsd", u"Package", false, false, nullptr);
    
    // Insert a link to the file in the local file system and display it as an icon.
    builder->InsertOleObject(get_ImageDir() + u"Microsoft Visio drawing.vsd", u"Package", true, true, nullptr);
    
    // Inserting OLE objects creates shapes that store these objects.
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Drawing::Shape>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToArray();
    
    ASSERT_EQ(2, shapes->get_Length());
    ASSERT_EQ(2, shapes->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Drawing::Shape>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Drawing::Shape> s)>>([](System::SharedPtr<Aspose::Words::Drawing::Shape> s) -> bool
    {
        return s->get_ShapeType() == Aspose::Words::Drawing::ShapeType::OleObject;
    }))));
    
    // If a shape contains an OLE object, it will have a valid "OleFormat" property,
    // which we can use to verify some aspects of the shape.
    System::SharedPtr<Aspose::Words::Drawing::OleFormat> oleFormat = shapes[0]->get_OleFormat();
    
    ASPOSE_ASSERT_EQ(false, oleFormat->get_IsLink());
    ASPOSE_ASSERT_EQ(false, oleFormat->get_OleIcon());
    
    oleFormat = shapes[1]->get_OleFormat();
    
    ASPOSE_ASSERT_EQ(true, oleFormat->get_IsLink());
    ASPOSE_ASSERT_EQ(true, oleFormat->get_OleIcon());
    
    ASSERT_TRUE(oleFormat->get_SourceFullName().EndsWith(System::String(u"Images") + System::IO::Path::DirectorySeparatorChar + u"Microsoft Visio drawing.vsd"));
    ASSERT_EQ(u"", oleFormat->get_SourceItem());
    
    ASSERT_EQ(u"Microsoft Visio drawing.vsd", oleFormat->get_IconCaption());
    
    doc->Save(get_ArtifactsDir() + u"Shape.OleLinks.docx");
    
    // If the object contains OLE data, we can access it using a stream.
    {
        System::SharedPtr<System::IO::MemoryStream> stream = oleFormat->GetOleEntry(u"\x0001" u"CompObj");
        System::ArrayPtr<uint8_t> oleEntryBytes = stream->ToArray();
        ASSERT_EQ(76, oleEntryBytes->get_Length());
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, OleLinks)
{
    s_instance->OleLinks();
}

} // namespace gtest_test

void ExShape::OleControlCollection()
{
    //ExStart
    //ExFor:OleFormat.Clsid
    //ExFor:Forms2OleControlCollection
    //ExFor:Forms2OleControlCollection.Count
    //ExFor:Forms2OleControlCollection.Item(Int32)
    //ExSummary:Shows how to access an OLE control embedded in a document and its child controls.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"OLE ActiveX controls.docm");
    
    // Shapes store and display OLE objects in the document's body.
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(u"6e182020-f460-11ce-9bcd-00aa00608e01", System::ObjectExt::ToString(shape->get_OleFormat()->get_Clsid()));
    
    auto oleControl = System::ExplicitCast<Aspose::Words::Drawing::Ole::Forms2OleControl>(shape->get_OleFormat()->get_OleControl());
    
    // Some OLE controls may contain child controls, such as the one in this document with three options buttons.
    System::SharedPtr<Aspose::Words::Drawing::Ole::Forms2OleControlCollection> oleControlCollection = oleControl->get_ChildNodes();
    
    ASSERT_EQ(3, oleControlCollection->get_Count());
    
    ASSERT_EQ(u"C#", oleControlCollection->idx_get(0)->get_Caption());
    ASSERT_EQ(u"1", oleControlCollection->idx_get(0)->get_Value());
    
    ASSERT_EQ(u"Visual Basic", oleControlCollection->idx_get(1)->get_Caption());
    ASSERT_EQ(u"0", oleControlCollection->idx_get(1)->get_Value());
    
    ASSERT_EQ(u"Delphi", oleControlCollection->idx_get(2)->get_Caption());
    ASSERT_EQ(u"0", oleControlCollection->idx_get(2)->get_Value());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, OleControlCollection)
{
    s_instance->OleControlCollection();
}

} // namespace gtest_test

void ExShape::SuggestedFileName()
{
    //ExStart
    //ExFor:OleFormat.SuggestedFileName
    //ExSummary:Shows how to get an OLE object's suggested file name.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"OLE shape.rtf");
    
    auto oleShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->get_FirstSection()->get_Body()->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    // OLE objects can provide a suggested filename and extension,
    // which we can use when saving the object's contents into a file in the local file system.
    System::String suggestedFileName = oleShape->get_OleFormat()->get_SuggestedFileName();
    
    ASSERT_EQ(u"CSV.csv", suggestedFileName);
    
    {
        auto fileStream = System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + suggestedFileName, System::IO::FileMode::Create);
        oleShape->get_OleFormat()->Save(fileStream);
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, SuggestedFileName)
{
    s_instance->SuggestedFileName();
}

} // namespace gtest_test

void ExShape::ObjectDidNotHaveSuggestedFileName()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"ActiveX controls.docx");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    ASSERT_EQ(System::String::Empty, shape->get_OleFormat()->get_SuggestedFileName());
}

namespace gtest_test
{

TEST_F(ExShape, ObjectDidNotHaveSuggestedFileName)
{
    s_instance->ObjectDidNotHaveSuggestedFileName();
}

} // namespace gtest_test

void ExShape::RenderOfficeMath()
{
    //ExStart
    //ExFor:ImageSaveOptions.Scale
    //ExFor:OfficeMath.GetMathRenderer
    //ExFor:NodeRendererBase.Save(String, ImageSaveOptions)
    //ExSummary:Shows how to render an Office Math object into an image file in the local file system.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto math = System::ExplicitCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    
    // Create an "ImageSaveOptions" object to pass to the node renderer's "Save" method to modify
    // how it renders the OfficeMath node into an image.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
    
    // Set the "Scale" property to 5 to render the object to five times its original size.
    saveOptions->set_Scale(5.0f);
    
    math->GetMathRenderer()->Save(get_ArtifactsDir() + u"Shape.RenderOfficeMath.png", saveOptions);
    //ExEnd
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(813, 87, get_ArtifactsDir() + u"Shape.RenderOfficeMath.png");
}

namespace gtest_test
{

TEST_F(ExShape, RenderOfficeMath)
{
    s_instance->RenderOfficeMath();
}

} // namespace gtest_test

void ExShape::OfficeMathDisplayException()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto officeMath = System::ExplicitCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    officeMath->set_DisplayType(Aspose::Words::Math::OfficeMathDisplayType::Display);
    
    ASSERT_THROW(static_cast<std::function<void()>>([&officeMath]() -> void
    {
        officeMath->set_Justification(Aspose::Words::Math::OfficeMathJustification::Inline);
    })(), System::ArgumentException);
}

namespace gtest_test
{

TEST_F(ExShape, OfficeMathDisplayException)
{
    s_instance->OfficeMathDisplayException();
}

} // namespace gtest_test

void ExShape::OfficeMathDefaultValue()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto officeMath = System::ExplicitCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 6, true));
    
    ASSERT_EQ(Aspose::Words::Math::OfficeMathDisplayType::Inline, officeMath->get_DisplayType());
    ASSERT_EQ(Aspose::Words::Math::OfficeMathJustification::Inline, officeMath->get_Justification());
}

namespace gtest_test
{

TEST_F(ExShape, OfficeMathDefaultValue)
{
    s_instance->OfficeMathDefaultValue();
}

} // namespace gtest_test

void ExShape::OfficeMath()
{
    //ExStart
    //ExFor:OfficeMath
    //ExFor:OfficeMath.DisplayType
    //ExFor:OfficeMath.Justification
    //ExFor:OfficeMath.NodeType
    //ExFor:OfficeMath.ParentParagraph
    //ExFor:OfficeMathDisplayType
    //ExFor:OfficeMathJustification
    //ExSummary:Shows how to set office math display formatting.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto officeMath = System::ExplicitCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    
    // OfficeMath nodes that are children of other OfficeMath nodes are always inline.
    // The node we are working with is the base node to change its location and display type.
    ASSERT_EQ(Aspose::Words::Math::MathObjectType::OMathPara, officeMath->get_MathObjectType());
    ASSERT_EQ(Aspose::Words::NodeType::OfficeMath, officeMath->get_NodeType());
    ASPOSE_ASSERT_EQ(officeMath->get_ParentNode(), officeMath->get_ParentParagraph());
    
    // Change the location and display type of the OfficeMath node.
    officeMath->set_DisplayType(Aspose::Words::Math::OfficeMathDisplayType::Display);
    officeMath->set_Justification(Aspose::Words::Math::OfficeMathJustification::Left);
    
    doc->Save(get_ArtifactsDir() + u"Shape.OfficeMath.docx");
    //ExEnd
    
    ASSERT_TRUE(Aspose::Words::ApiExamples::DocumentHelper::CompareDocs(get_ArtifactsDir() + u"Shape.OfficeMath.docx", get_GoldsDir() + u"Shape.OfficeMath Gold.docx"));
}

namespace gtest_test
{

TEST_F(ExShape, OfficeMath)
{
    s_instance->OfficeMath();
}

} // namespace gtest_test

void ExShape::CannotBeSetDisplayWithInlineJustification()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto officeMath = System::ExplicitCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    officeMath->set_DisplayType(Aspose::Words::Math::OfficeMathDisplayType::Display);
    
    ASSERT_THROW(static_cast<std::function<void()>>([&officeMath]() -> void
    {
        officeMath->set_Justification(Aspose::Words::Math::OfficeMathJustification::Inline);
    })(), System::ArgumentException);
}

namespace gtest_test
{

TEST_F(ExShape, CannotBeSetDisplayWithInlineJustification)
{
    s_instance->CannotBeSetDisplayWithInlineJustification();
}

} // namespace gtest_test

void ExShape::CannotBeSetInlineDisplayWithJustification()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto officeMath = System::ExplicitCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    officeMath->set_DisplayType(Aspose::Words::Math::OfficeMathDisplayType::Inline);
    
    ASSERT_THROW(static_cast<std::function<void()>>([&officeMath]() -> void
    {
        officeMath->set_Justification(Aspose::Words::Math::OfficeMathJustification::Center);
    })(), System::ArgumentException);
}

namespace gtest_test
{

TEST_F(ExShape, CannotBeSetInlineDisplayWithJustification)
{
    s_instance->CannotBeSetInlineDisplayWithJustification();
}

} // namespace gtest_test

void ExShape::OfficeMathDisplayNestedObjects()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto officeMath = System::ExplicitCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    
    ASSERT_EQ(Aspose::Words::Math::OfficeMathDisplayType::Display, officeMath->get_DisplayType());
    ASSERT_EQ(Aspose::Words::Math::OfficeMathJustification::Center, officeMath->get_Justification());
}

namespace gtest_test
{

TEST_F(ExShape, OfficeMathDisplayNestedObjects)
{
    s_instance->OfficeMathDisplayNestedObjects();
}

} // namespace gtest_test

void ExShape::WorkWithMathObjectType(int32_t index, Aspose::Words::Math::MathObjectType objectType)
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto officeMath = System::ExplicitCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, index, true));
    ASSERT_EQ(objectType, officeMath->get_MathObjectType());
}

namespace gtest_test
{

using ExShape_WorkWithMathObjectType_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExShape::WorkWithMathObjectType)>::type;

struct ExShape_WorkWithMathObjectType : public ExShape, public Aspose::Words::ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_WorkWithMathObjectType_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(0, Aspose::Words::Math::MathObjectType::OMathPara),
            std::make_tuple(1, Aspose::Words::Math::MathObjectType::OMath),
            std::make_tuple(2, Aspose::Words::Math::MathObjectType::Supercript),
            std::make_tuple(3, Aspose::Words::Math::MathObjectType::Argument),
            std::make_tuple(4, Aspose::Words::Math::MathObjectType::SuperscriptPart),
        };
    }
};

TEST_P(ExShape_WorkWithMathObjectType, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->WorkWithMathObjectType(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_WorkWithMathObjectType, ::testing::ValuesIn(ExShape_WorkWithMathObjectType::TestCases()));

} // namespace gtest_test

void ExShape::AspectRatio(bool lockAspectRatio)
{
    //ExStart
    //ExFor:ShapeBase.AspectRatioLocked
    //ExSummary:Shows how to lock/unlock a shape's aspect ratio.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a shape. If we open this document in Microsoft Word, we can left click the shape to reveal
    // eight sizing handles around its perimeter, which we can click and drag to change its size.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    // Set the "AspectRatioLocked" property to "true" to preserve the shape's aspect ratio
    // when using any of the four diagonal sizing handles, which change both the image's height and width.
    // Using any orthogonal sizing handles that either change the height or width will still change the aspect ratio.
    // Set the "AspectRatioLocked" property to "false" to allow us to
    // freely change the image's aspect ratio with all sizing handles.
    shape->set_AspectRatioLocked(lockAspectRatio);
    
    doc->Save(get_ArtifactsDir() + u"Shape.AspectRatio.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.AspectRatio.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASPOSE_ASSERT_EQ(lockAspectRatio, shape->get_AspectRatioLocked());
}

namespace gtest_test
{

using ExShape_AspectRatio_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExShape::AspectRatio)>::type;

struct ExShape_AspectRatio : public ExShape, public Aspose::Words::ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_AspectRatio_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
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

} // namespace gtest_test

void ExShape::MarkupLanguageByDefault()
{
    //ExStart
    //ExFor:ShapeBase.MarkupLanguage
    //ExFor:ShapeBase.SizeInPoints
    //ExSummary:Shows how to verify a shape's size and markup language.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertImage(get_ImageDir() + u"Transparent background logo.png");
    
    ASSERT_EQ(Aspose::Words::Drawing::ShapeMarkupLanguage::Dml, shape->get_MarkupLanguage());
    ASPOSE_ASSERT_EQ(System::Drawing::SizeF(300.0f, 300.0f), shape->get_SizeInPoints());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, MarkupLanguageByDefault)
{
    s_instance->MarkupLanguageByDefault();
}

} // namespace gtest_test

void ExShape::MarkupLanguageForDifferentMsWordVersions(Aspose::Words::Settings::MsWordVersion msWordVersion, Aspose::Words::Drawing::ShapeMarkupLanguage shapeMarkupLanguage)
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->get_CompatibilityOptions()->OptimizeFor(msWordVersion);
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->InsertImage(get_ImageDir() + u"Transparent background logo.png");
    
    for (auto&& shape : System::IterateOver(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()))
    {
        ASSERT_EQ(shapeMarkupLanguage, shape->get_MarkupLanguage());
    }
}

namespace gtest_test
{

using ExShape_MarkupLanguageForDifferentMsWordVersions_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExShape::MarkupLanguageForDifferentMsWordVersions)>::type;

struct ExShape_MarkupLanguageForDifferentMsWordVersions : public ExShape, public Aspose::Words::ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_MarkupLanguageForDifferentMsWordVersions_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Settings::MsWordVersion::Word2000, Aspose::Words::Drawing::ShapeMarkupLanguage::Vml),
            std::make_tuple(Aspose::Words::Settings::MsWordVersion::Word2002, Aspose::Words::Drawing::ShapeMarkupLanguage::Vml),
            std::make_tuple(Aspose::Words::Settings::MsWordVersion::Word2003, Aspose::Words::Drawing::ShapeMarkupLanguage::Vml),
            std::make_tuple(Aspose::Words::Settings::MsWordVersion::Word2007, Aspose::Words::Drawing::ShapeMarkupLanguage::Vml),
            std::make_tuple(Aspose::Words::Settings::MsWordVersion::Word2010, Aspose::Words::Drawing::ShapeMarkupLanguage::Dml),
            std::make_tuple(Aspose::Words::Settings::MsWordVersion::Word2013, Aspose::Words::Drawing::ShapeMarkupLanguage::Dml),
            std::make_tuple(Aspose::Words::Settings::MsWordVersion::Word2016, Aspose::Words::Drawing::ShapeMarkupLanguage::Dml),
        };
    }
};

TEST_P(ExShape_MarkupLanguageForDifferentMsWordVersions, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MarkupLanguageForDifferentMsWordVersions(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_MarkupLanguageForDifferentMsWordVersions, ::testing::ValuesIn(ExShape_MarkupLanguageForDifferentMsWordVersions::TestCases()));

} // namespace gtest_test

void ExShape::Stroke()
{
    //ExStart
    //ExFor:Stroke
    //ExFor:Stroke.On
    //ExFor:Stroke.Weight
    //ExFor:Stroke.JoinStyle
    //ExFor:Stroke.LineStyle
    //ExFor:Stroke.Fill
    //ExFor:ShapeLineStyle
    //ExSummary:Shows how change stroke properties.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 100, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 100, 200, 200, Aspose::Words::Drawing::WrapType::None);
    
    // Basic shapes, such as the rectangle, have two visible parts.
    // 1 -  The fill, which applies to the area within the outline of the shape:
    shape->get_Fill()->set_ForeColor(System::Drawing::Color::get_White());
    
    // 2 -  The stroke, which marks the outline of the shape:
    // Modify various properties of this shape's stroke.
    System::SharedPtr<Aspose::Words::Drawing::Stroke> stroke = shape->get_Stroke();
    stroke->set_On(true);
    stroke->set_Weight(5);
    stroke->set_Color(System::Drawing::Color::get_Red());
    stroke->set_DashStyle(Aspose::Words::Drawing::DashStyle::ShortDashDotDot);
    stroke->set_JoinStyle(Aspose::Words::Drawing::JoinStyle::Miter);
    stroke->set_EndCap(Aspose::Words::Drawing::EndCap::Square);
    stroke->set_LineStyle(Aspose::Words::Drawing::ShapeLineStyle::Triple);
    stroke->get_Fill()->TwoColorGradient(System::Drawing::Color::get_Red(), System::Drawing::Color::get_Blue(), Aspose::Words::Drawing::GradientStyle::Vertical, Aspose::Words::Drawing::GradientVariant::Variant1);
    
    doc->Save(get_ArtifactsDir() + u"Shape.Stroke.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.Stroke.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    stroke = shape->get_Stroke();
    
    ASPOSE_ASSERT_EQ(true, stroke->get_On());
    ASPOSE_ASSERT_EQ(5, stroke->get_Weight());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), stroke->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::Drawing::DashStyle::ShortDashDotDot, stroke->get_DashStyle());
    ASSERT_EQ(Aspose::Words::Drawing::JoinStyle::Miter, stroke->get_JoinStyle());
    ASSERT_EQ(Aspose::Words::Drawing::EndCap::Square, stroke->get_EndCap());
    ASSERT_EQ(Aspose::Words::Drawing::ShapeLineStyle::Triple, stroke->get_LineStyle());
}

namespace gtest_test
{

TEST_F(ExShape, Stroke)
{
    s_instance->Stroke();
}

} // namespace gtest_test

void ExShape::InsertOleObjectAsHtmlFile()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertOleObject(u"http://www.aspose.com", u"htmlfile", true, false, nullptr);
    
    doc->Save(get_ArtifactsDir() + u"Shape.InsertOleObjectAsHtmlFile.docx");
}

namespace gtest_test
{

TEST_F(ExShape, InsertOleObjectAsHtmlFile)
{
    s_instance->InsertOleObjectAsHtmlFile();
}

} // namespace gtest_test

void ExShape::InsertOlePackage()
{
    //ExStart
    //ExFor:OlePackage
    //ExFor:OleFormat.OlePackage
    //ExFor:OlePackage.FileName
    //ExFor:OlePackage.DisplayName
    //ExSummary:Shows how insert an OLE object into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // OLE objects allow us to open other files in the local file system using another installed application
    // in our operating system by double-clicking on the shape that contains the OLE object in the document body.
    // In this case, our external file will be a ZIP archive.
    System::ArrayPtr<uint8_t> zipFileBytes = System::IO::File::ReadAllBytes(get_DatabaseDir() + u"cat001.zip");
    
    {
        auto stream = System::MakeObject<System::IO::MemoryStream>(zipFileBytes);
        System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertOleObject(stream, u"Package", true, nullptr);
        
        shape->get_OleFormat()->get_OlePackage()->set_FileName(u"Package file name.zip");
        shape->get_OleFormat()->get_OlePackage()->set_DisplayName(u"Package display name.zip");
    }
    
    doc->Save(get_ArtifactsDir() + u"Shape.InsertOlePackage.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.InsertOlePackage.docx");
    auto getShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(u"Package file name.zip", getShape->get_OleFormat()->get_OlePackage()->get_FileName());
    ASSERT_EQ(u"Package display name.zip", getShape->get_OleFormat()->get_OlePackage()->get_DisplayName());
}

namespace gtest_test
{

TEST_F(ExShape, InsertOlePackage)
{
    s_instance->InsertOlePackage();
}

} // namespace gtest_test

void ExShape::GetAccessToOlePackage()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> oleObject = builder->InsertOleObject(get_MyDir() + u"Spreadsheet.xlsx", false, false, nullptr);
    System::SharedPtr<Aspose::Words::Drawing::Shape> oleObjectAsOlePackage = builder->InsertOleObject(get_MyDir() + u"Spreadsheet.xlsx", u"Excel.Sheet", false, false, nullptr);
    
    ASPOSE_ASSERT_EQ(nullptr, oleObject->get_OleFormat()->get_OlePackage());
    ASPOSE_ASSERT_EQ(System::ObjectExt::GetType<Aspose::Words::Drawing::OlePackage>(), System::ObjectExt::GetType(oleObjectAsOlePackage->get_OleFormat()->get_OlePackage()));
}

namespace gtest_test
{

TEST_F(ExShape, GetAccessToOlePackage)
{
    s_instance->GetAccessToOlePackage();
}

} // namespace gtest_test

void ExShape::Resize()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 200, 300);
    shape->set_Height(300);
    shape->set_Width(500);
    shape->set_Rotation(30);
    
    doc->Save(get_ArtifactsDir() + u"Shape.Resize.docx");
}

namespace gtest_test
{

TEST_F(ExShape, Resize)
{
    s_instance->Resize();
}

} // namespace gtest_test

void ExShape::Calendar()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->StartTable();
    builder->get_RowFormat()->set_Height(100);
    builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Exactly);
    
    for (int32_t i = 0; i < 31; i++)
    {
        if (i != 0 && i % 7 == 0)
        {
            builder->EndRow();
        }
        builder->InsertCell();
        builder->Write(u"Cell contents");
    }
    
    builder->EndTable();
    
    System::SharedPtr<Aspose::Words::NodeCollection> runs = doc->GetChildNodes(Aspose::Words::NodeType::Run, true);
    int32_t num = 1;
    
    for (auto&& run : System::IterateOver(runs->LINQ_OfType<System::SharedPtr<Aspose::Words::Run> >()))
    {
        auto watermark = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::TextPlainText);
        watermark->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::Page);
        watermark->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::Page);
        watermark->set_Width(30);
        watermark->set_Height(30);
        watermark->set_HorizontalAlignment(Aspose::Words::Drawing::HorizontalAlignment::Center);
        watermark->set_VerticalAlignment(Aspose::Words::Drawing::VerticalAlignment::Center);
        watermark->set_Rotation(-40);
        
        
        watermark->get_Fill()->set_ForeColor(System::Drawing::Color::get_Gainsboro());
        watermark->set_StrokeColor(System::Drawing::Color::get_Gainsboro());
        
        watermark->get_TextPath()->set_Text(System::String::Format(u"{0}", num));
        watermark->get_TextPath()->set_FontFamily(u"Arial");
        
        watermark->set_Name(System::String::Format(u"Watermark_{0}", num++));
        
        watermark->set_BehindText(true);
        
        builder->MoveTo(run);
        builder->InsertNode(watermark);
    }
    
    doc->Save(get_ArtifactsDir() + u"Shape.Calendar.docx");
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.Calendar.docx");
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Drawing::Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_Cast<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToList();
    
    ASSERT_EQ(31, shapes->get_Count());
    
    for (auto&& shape : shapes)
    {
        Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, System::String::Format(u"Watermark_{0}", shapes->IndexOf(shape) + 1), 30.0, 30.0, 0.0, 0.0, shape);
    }
}

namespace gtest_test
{

TEST_F(ExShape, Calendar)
{
    s_instance->Calendar();
}

} // namespace gtest_test

void ExShape::IsLayoutInCell(bool isLayoutInCell)
{
    //ExStart
    //ExFor:ShapeBase.IsLayoutInCell
    //ExSummary:Shows how to determine how to display a shape in a table cell.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->InsertCell();
    builder->EndTable();
    
    auto tableStyle = System::ExplicitCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle1"));
    tableStyle->set_BottomPadding(20);
    tableStyle->set_LeftPadding(10);
    tableStyle->set_RightPadding(10);
    tableStyle->set_TopPadding(20);
    tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Black());
    tableStyle->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::Single);
    
    table->set_Style(tableStyle);
    
    builder->MoveTo(table->get_FirstRow()->get_FirstCell()->get_FirstParagraph());
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, Aspose::Words::Drawing::RelativeHorizontalPosition::LeftMargin, 50, Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin, 100, 100, 100, Aspose::Words::Drawing::WrapType::None);
    
    // Set the "IsLayoutInCell" property to "true" to display the shape as an inline element inside the cell's paragraph.
    // The coordinate origin that will determine the shape's location will be the top left corner of the shape's cell.
    // If we re-size the cell, the shape will move to maintain the same position starting from the cell's top left.
    // Set the "IsLayoutInCell" property to "false" to display the shape as an independent floating shape.
    // The coordinate origin that will determine the shape's location will be the top left corner of the page,
    // and the shape will not respond to any re-sizing of its cell.
    shape->set_IsLayoutInCell(isLayoutInCell);
    
    // We can only apply the "IsLayoutInCell" property to floating shapes.
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    
    doc->Save(get_ArtifactsDir() + u"Shape.LayoutInTableCell.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.LayoutInTableCell.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(table->get_FirstRow()->get_FirstCell()->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASPOSE_ASSERT_EQ(isLayoutInCell, shape->get_IsLayoutInCell());
}

namespace gtest_test
{

using ExShape_IsLayoutInCell_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExShape::IsLayoutInCell)>::type;

struct ExShape_IsLayoutInCell : public ExShape, public Aspose::Words::ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_IsLayoutInCell_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
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

} // namespace gtest_test

void ExShape::ShapeInsertion()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertShape(ShapeType, RelativeHorizontalPosition, double, RelativeVerticalPosition, double, double, double, WrapType)
    //ExFor:DocumentBuilder.InsertShape(ShapeType, double, double)
    //ExFor:OoxmlCompliance
    //ExFor:OoxmlSaveOptions.Compliance
    //ExSummary:Shows how to insert DML shapes into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are two wrapping types that shapes may have.
    // 1 -  Floating:
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::TopCornersRounded, Aspose::Words::Drawing::RelativeHorizontalPosition::Page, 100, Aspose::Words::Drawing::RelativeVerticalPosition::Page, 100, 50, 50, Aspose::Words::Drawing::WrapType::None);
    
    // 2 -  Inline:
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::DiagonalCornersRounded, 50, 50);
    
    // If you need to create "non-primitive" shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
    // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, or DiagonalCornersRounded,
    // then save the document with "Strict" or "Transitional" compliance, which allows saving shape as DML.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>(Aspose::Words::SaveFormat::Docx);
    saveOptions->set_Compliance(Aspose::Words::Saving::OoxmlCompliance::Iso29500_2008_Transitional);
    
    doc->Save(get_ArtifactsDir() + u"Shape.ShapeInsertion.docx", saveOptions);
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.ShapeInsertion.docx");
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Drawing::Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_Cast<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToList();
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TopCornersRounded, u"TopCornersRounded 100002", 50.0, 50.0, 100.0, 100.0, shapes->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::DiagonalCornersRounded, u"DiagonalCornersRounded 100004", 50.0, 50.0, 0.0, 0.0, shapes->idx_get(1));
}

namespace gtest_test
{

TEST_F(ExShape, ShapeInsertion)
{
    s_instance->ShapeInsertion();
}

} // namespace gtest_test

void ExShape::VisitShapes()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Revision shape.docx");
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    //ExSkip
    
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExShape::ShapeAppearancePrinter>();
    doc->Accept(visitor);
    
    std::cout << visitor->GetText() << std::endl;
}

namespace gtest_test
{

TEST_F(ExShape, VisitShapes)
{
    s_instance->VisitShapes();
}

} // namespace gtest_test

void ExShape::SignatureLine()
{
    //ExStart
    //ExFor:Shape.SignatureLine
    //ExFor:ShapeBase.IsSignatureLine
    //ExFor:SignatureLine
    //ExFor:SignatureLine.AllowComments
    //ExFor:SignatureLine.DefaultInstructions
    //ExFor:SignatureLine.Email
    //ExFor:SignatureLine.Instructions
    //ExFor:SignatureLine.ShowDate
    //ExFor:SignatureLine.Signer
    //ExFor:SignatureLine.SignerTitle
    //ExSummary:Shows how to create a line for a signature and insert it into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto options = System::MakeObject<Aspose::Words::SignatureLineOptions>();
    options->set_AllowComments(true);
    options->set_DefaultInstructions(true);
    options->set_Email(u"john.doe@management.com");
    options->set_Instructions(u"Please sign here");
    options->set_ShowDate(true);
    options->set_Signer(u"John Doe");
    options->set_SignerTitle(u"Senior Manager");
    
    // Insert a shape that will contain a signature line, whose appearance we will
    // customize using the "SignatureLineOptions" object we have created above.
    // If we insert a shape whose coordinates originate at the bottom right hand corner of the page,
    // we will need to supply negative x and y coordinates to bring the shape into view.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertSignatureLine(options, Aspose::Words::Drawing::RelativeHorizontalPosition::RightMargin, -170.0, Aspose::Words::Drawing::RelativeVerticalPosition::BottomMargin, -60.0, Aspose::Words::Drawing::WrapType::None);
    
    ASSERT_TRUE(shape->get_IsSignatureLine());
    
    // Verify the properties of our signature line via its Shape object.
    System::SharedPtr<Aspose::Words::Drawing::SignatureLine> signatureLine = shape->get_SignatureLine();
    
    ASSERT_EQ(u"john.doe@management.com", signatureLine->get_Email());
    ASSERT_EQ(u"John Doe", signatureLine->get_Signer());
    ASSERT_EQ(u"Senior Manager", signatureLine->get_SignerTitle());
    ASSERT_EQ(u"Please sign here", signatureLine->get_Instructions());
    ASSERT_TRUE(signatureLine->get_ShowDate());
    ASSERT_TRUE(signatureLine->get_AllowComments());
    ASSERT_TRUE(signatureLine->get_DefaultInstructions());
    
    doc->Save(get_ArtifactsDir() + u"Shape.SignatureLine.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.SignatureLine.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Image, System::String::Empty, 192.75, 96.75, -60.0, -170.0, shape);
    ASSERT_TRUE(shape->get_IsSignatureLine());
    
    signatureLine = shape->get_SignatureLine();
    
    ASSERT_EQ(u"john.doe@management.com", signatureLine->get_Email());
    ASSERT_EQ(u"John Doe", signatureLine->get_Signer());
    ASSERT_EQ(u"Senior Manager", signatureLine->get_SignerTitle());
    ASSERT_EQ(u"Please sign here", signatureLine->get_Instructions());
    ASSERT_TRUE(signatureLine->get_ShowDate());
    ASSERT_TRUE(signatureLine->get_AllowComments());
    ASSERT_TRUE(signatureLine->get_DefaultInstructions());
    ASSERT_FALSE(signatureLine->get_IsSigned());
    ASSERT_FALSE(signatureLine->get_IsValid());
}

namespace gtest_test
{

TEST_F(ExShape, SignatureLine)
{
    s_instance->SignatureLine();
}

} // namespace gtest_test

void ExShape::TextBoxLayoutFlow(Aspose::Words::Drawing::LayoutFlow layoutFlow)
{
    //ExStart
    //ExFor:Shape.TextBox
    //ExFor:Shape.LastParagraph
    //ExFor:TextBox
    //ExFor:TextBox.LayoutFlow
    //ExSummary:Shows how to set the orientation of text inside a text box.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBoxShape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 150, 100);
    System::SharedPtr<Aspose::Words::Drawing::TextBox> textBox = textBoxShape->get_TextBox();
    
    // Move the document builder to inside the TextBox and add text.
    builder->MoveTo(textBoxShape->get_LastParagraph());
    builder->Writeln(u"Hello world!");
    builder->Write(u"Hello again!");
    
    // Set the "LayoutFlow" property to set an orientation for the text contents of this text box.
    textBox->set_LayoutFlow(layoutFlow);
    
    doc->Save(get_ArtifactsDir() + u"Shape.TextBoxLayoutFlow.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.TextBoxLayoutFlow.docx");
    textBoxShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100002", 150.0, 100.0, 0.0, 0.0, textBoxShape);
    
    Aspose::Words::Drawing::LayoutFlow expectedLayoutFlow;
    
    switch (layoutFlow)
    {
        case Aspose::Words::Drawing::LayoutFlow::BottomToTop:
        case Aspose::Words::Drawing::LayoutFlow::Horizontal:
        case Aspose::Words::Drawing::LayoutFlow::TopToBottomIdeographic:
        case Aspose::Words::Drawing::LayoutFlow::Vertical:
            expectedLayoutFlow = layoutFlow;
            break;
        
        case Aspose::Words::Drawing::LayoutFlow::TopToBottom:
            expectedLayoutFlow = Aspose::Words::Drawing::LayoutFlow::Vertical;
            break;
        
        default: 
            expectedLayoutFlow = Aspose::Words::Drawing::LayoutFlow::Horizontal;
            break;
        
    }
    
    Aspose::Words::ApiExamples::TestUtil::VerifyTextBox(expectedLayoutFlow, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, textBoxShape->get_TextBox());
    ASSERT_EQ(u"Hello world!\rHello again!", textBoxShape->GetText().Trim());
}

namespace gtest_test
{

using ExShape_TextBoxLayoutFlow_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExShape::TextBoxLayoutFlow)>::type;

struct ExShape_TextBoxLayoutFlow : public ExShape, public Aspose::Words::ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_TextBoxLayoutFlow_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Drawing::LayoutFlow::Vertical),
            std::make_tuple(Aspose::Words::Drawing::LayoutFlow::Horizontal),
            std::make_tuple(Aspose::Words::Drawing::LayoutFlow::HorizontalIdeographic),
            std::make_tuple(Aspose::Words::Drawing::LayoutFlow::BottomToTop),
            std::make_tuple(Aspose::Words::Drawing::LayoutFlow::TopToBottom),
            std::make_tuple(Aspose::Words::Drawing::LayoutFlow::TopToBottomIdeographic),
        };
    }
};

TEST_P(ExShape_TextBoxLayoutFlow, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TextBoxLayoutFlow(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_TextBoxLayoutFlow, ::testing::ValuesIn(ExShape_TextBoxLayoutFlow::TestCases()));

} // namespace gtest_test

void ExShape::TextBoxFitShapeToText()
{
    //ExStart
    //ExFor:TextBox
    //ExFor:TextBox.FitShapeToText
    //ExSummary:Shows how to get a text box to resize itself to fit its contents tightly.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBoxShape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 150, 100);
    System::SharedPtr<Aspose::Words::Drawing::TextBox> textBox = textBoxShape->get_TextBox();
    
    // Apply these values to both these members to get the parent shape to fit
    // tightly around the text contents, ignoring the dimensions we have set.
    textBox->set_FitShapeToText(true);
    textBox->set_TextBoxWrapMode(Aspose::Words::Drawing::TextBoxWrapMode::None);
    
    builder->MoveTo(textBoxShape->get_LastParagraph());
    builder->Write(u"Text fit tightly inside textbox.");
    
    doc->Save(get_ArtifactsDir() + u"Shape.TextBoxFitShapeToText.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.TextBoxFitShapeToText.docx");
    textBoxShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100002", 150.0, 100.0, 0.0, 0.0, textBoxShape);
    Aspose::Words::ApiExamples::TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, true, Aspose::Words::Drawing::TextBoxWrapMode::None, 3.6, 3.6, 7.2, 7.2, textBoxShape->get_TextBox());
    ASSERT_EQ(u"Text fit tightly inside textbox.", textBoxShape->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExShape, TextBoxFitShapeToText)
{
    s_instance->TextBoxFitShapeToText();
}

} // namespace gtest_test

void ExShape::TextBoxMargins()
{
    //ExStart
    //ExFor:TextBox
    //ExFor:TextBox.InternalMarginBottom
    //ExFor:TextBox.InternalMarginLeft
    //ExFor:TextBox.InternalMarginRight
    //ExFor:TextBox.InternalMarginTop
    //ExSummary:Shows how to set internal margins for a text box.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert another textbox with specific margins.
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBoxShape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 100);
    System::SharedPtr<Aspose::Words::Drawing::TextBox> textBox = textBoxShape->get_TextBox();
    textBox->set_InternalMarginTop(15);
    textBox->set_InternalMarginBottom(15);
    textBox->set_InternalMarginLeft(15);
    textBox->set_InternalMarginRight(15);
    
    builder->MoveTo(textBoxShape->get_LastParagraph());
    builder->Write(u"Text placed according to textbox margins.");
    
    doc->Save(get_ArtifactsDir() + u"Shape.TextBoxMargins.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.TextBoxMargins.docx");
    textBoxShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100002", 100.0, 100.0, 0.0, 0.0, textBoxShape);
    Aspose::Words::ApiExamples::TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 15.0, 15.0, 15.0, 15.0, textBoxShape->get_TextBox());
    ASSERT_EQ(u"Text placed according to textbox margins.", textBoxShape->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExShape, TextBoxMargins)
{
    s_instance->TextBoxMargins();
}

} // namespace gtest_test

void ExShape::TextBoxContentsWrapMode(Aspose::Words::Drawing::TextBoxWrapMode textBoxWrapMode)
{
    //ExStart
    //ExFor:TextBox.TextBoxWrapMode
    //ExFor:TextBoxWrapMode
    //ExSummary:Shows how to set a wrapping mode for the contents of a text box.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBoxShape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 300, 300);
    System::SharedPtr<Aspose::Words::Drawing::TextBox> textBox = textBoxShape->get_TextBox();
    
    // Set the "TextBoxWrapMode" property to "TextBoxWrapMode.None" to increase the text box's width
    // to accommodate text, should it be large enough.
    // Set the "TextBoxWrapMode" property to "TextBoxWrapMode.Square" to
    // wrap all text inside the text box, preserving its dimensions.
    textBox->set_TextBoxWrapMode(textBoxWrapMode);
    
    builder->MoveTo(textBoxShape->get_LastParagraph());
    builder->get_Font()->set_Size(32);
    builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    
    doc->Save(get_ArtifactsDir() + u"Shape.TextBoxContentsWrapMode.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.TextBoxContentsWrapMode.docx");
    textBoxShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100002", 300.0, 300.0, 0.0, 0.0, textBoxShape);
    Aspose::Words::ApiExamples::TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, textBoxWrapMode, 3.6, 3.6, 7.2, 7.2, textBoxShape->get_TextBox());
    ASSERT_EQ(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.", textBoxShape->GetText().Trim());
}

namespace gtest_test
{

using ExShape_TextBoxContentsWrapMode_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExShape::TextBoxContentsWrapMode)>::type;

struct ExShape_TextBoxContentsWrapMode : public ExShape, public Aspose::Words::ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_TextBoxContentsWrapMode_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Drawing::TextBoxWrapMode::None),
            std::make_tuple(Aspose::Words::Drawing::TextBoxWrapMode::Square),
        };
    }
};

TEST_P(ExShape_TextBoxContentsWrapMode, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TextBoxContentsWrapMode(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_TextBoxContentsWrapMode, ::testing::ValuesIn(ExShape_TextBoxContentsWrapMode::TestCases()));

} // namespace gtest_test

void ExShape::TextBoxShapeType()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Set compatibility options to correctly using of VerticalAnchor property.
    doc->get_CompatibilityOptions()->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2016);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBoxShape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 100);
    // Not all formats are compatible with this one.
    // For most of the incompatible formats, AW generated warnings on save, so use doc.WarningCallback to check it.
    textBoxShape->get_TextBox()->set_VerticalAnchor(Aspose::Words::Drawing::TextBoxAnchor::Bottom);
    
    builder->MoveTo(textBoxShape->get_LastParagraph());
    builder->Write(u"Text placed bottom");
    
    doc->Save(get_ArtifactsDir() + u"Shape.TextBoxShapeType.docx");
}

namespace gtest_test
{

TEST_F(ExShape, TextBoxShapeType)
{
    s_instance->TextBoxShapeType();
}

} // namespace gtest_test

void ExShape::CreateLinkBetweenTextBoxes()
{
    //ExStart
    //ExFor:TextBox.IsValidLinkTarget(TextBox)
    //ExFor:TextBox.Next
    //ExFor:TextBox.Previous
    //ExFor:TextBox.BreakForwardLink
    //ExSummary:Shows how to link text boxes.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBoxShape1 = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 100);
    System::SharedPtr<Aspose::Words::Drawing::TextBox> textBox1 = textBoxShape1->get_TextBox();
    builder->Writeln();
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBoxShape2 = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 100);
    System::SharedPtr<Aspose::Words::Drawing::TextBox> textBox2 = textBoxShape2->get_TextBox();
    builder->Writeln();
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBoxShape3 = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 100);
    System::SharedPtr<Aspose::Words::Drawing::TextBox> textBox3 = textBoxShape3->get_TextBox();
    builder->Writeln();
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBoxShape4 = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 100);
    System::SharedPtr<Aspose::Words::Drawing::TextBox> textBox4 = textBoxShape4->get_TextBox();
    
    // Create links between some of the text boxes.
    if (textBox1->IsValidLinkTarget(textBox2))
    {
        textBox1->set_Next(textBox2);
    }
    
    if (textBox2->IsValidLinkTarget(textBox3))
    {
        textBox2->set_Next(textBox3);
    }
    
    // Only an empty text box may have a link.
    ASSERT_TRUE(textBox3->IsValidLinkTarget(textBox4));
    
    builder->MoveTo(textBoxShape4->get_LastParagraph());
    builder->Write(u"Hello world!");
    
    ASSERT_FALSE(textBox3->IsValidLinkTarget(textBox4));
    
    if (textBox1->get_Next() != nullptr && textBox1->get_Previous() == nullptr)
    {
        std::cout << "This TextBox is the head of the sequence" << std::endl;
    }
    
    if (textBox2->get_Next() != nullptr && textBox2->get_Previous() != nullptr)
    {
        std::cout << "This TextBox is the middle of the sequence" << std::endl;
    }
    
    if (textBox3->get_Next() == nullptr && textBox3->get_Previous() != nullptr)
    {
        std::cout << "This TextBox is the tail of the sequence" << std::endl;
        
        // Break the forward link between textBox2 and textBox3, and then verify that they are no longer linked.
        textBox3->get_Previous()->BreakForwardLink();
        ASSERT_TRUE(textBox2->get_Next() == nullptr);
        ASSERT_TRUE(textBox3->get_Previous() == nullptr);
    }
    
    doc->Save(get_ArtifactsDir() + u"Shape.CreateLinkBetweenTextBoxes.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.CreateLinkBetweenTextBoxes.docx");
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Drawing::Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToList();
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100002", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(0)->get_TextBox());
    ASSERT_EQ(System::String::Empty, shapes->idx_get(0)->GetText().Trim());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100004", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(1));
    Aspose::Words::ApiExamples::TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(1)->get_TextBox());
    ASSERT_EQ(System::String::Empty, shapes->idx_get(1)->GetText().Trim());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, u"TextBox 100006", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(2));
    Aspose::Words::ApiExamples::TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(2)->get_TextBox());
    ASSERT_EQ(System::String::Empty, shapes->idx_get(2)->GetText().Trim());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100008", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(3));
    Aspose::Words::ApiExamples::TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(3)->get_TextBox());
    ASSERT_EQ(u"Hello world!", shapes->idx_get(3)->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExShape, CreateLinkBetweenTextBoxes)
{
    s_instance->CreateLinkBetweenTextBoxes();
}

} // namespace gtest_test

void ExShape::VerticalAnchor(Aspose::Words::Drawing::TextBoxAnchor verticalAnchor)
{
    //ExStart
    //ExFor:CompatibilityOptions
    //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
    //ExFor:TextBoxAnchor
    //ExFor:TextBox.VerticalAnchor
    //ExSummary:Shows how to vertically align the text contents of a text box.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 200, 200);
    
    // Set the "VerticalAnchor" property to "TextBoxAnchor.Top" to
    // align the text in this text box with the top side of the shape.
    // Set the "VerticalAnchor" property to "TextBoxAnchor.Middle" to
    // align the text in this text box to the center of the shape.
    // Set the "VerticalAnchor" property to "TextBoxAnchor.Bottom" to
    // align the text in this text box to the bottom of the shape.
    shape->get_TextBox()->set_VerticalAnchor(verticalAnchor);
    
    builder->MoveTo(shape->get_FirstParagraph());
    builder->Write(u"Hello world!");
    
    // The vertical aligning of text inside text boxes is available from Microsoft Word 2007 onwards.
    doc->get_CompatibilityOptions()->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2007);
    doc->Save(get_ArtifactsDir() + u"Shape.VerticalAnchor.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.VerticalAnchor.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100002", 200.0, 200.0, 0.0, 0.0, shape);
    Aspose::Words::ApiExamples::TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shape->get_TextBox());
    ASSERT_EQ(verticalAnchor, shape->get_TextBox()->get_VerticalAnchor());
    ASSERT_EQ(u"Hello world!", shape->GetText().Trim());
}

namespace gtest_test
{

using ExShape_VerticalAnchor_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExShape::VerticalAnchor)>::type;

struct ExShape_VerticalAnchor : public ExShape, public Aspose::Words::ApiExamples::ExShape, public ::testing::WithParamInterface<ExShape_VerticalAnchor_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Drawing::TextBoxAnchor::Top),
            std::make_tuple(Aspose::Words::Drawing::TextBoxAnchor::Middle),
            std::make_tuple(Aspose::Words::Drawing::TextBoxAnchor::Bottom),
        };
    }
};

TEST_P(ExShape_VerticalAnchor, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->VerticalAnchor(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExShape_VerticalAnchor, ::testing::ValuesIn(ExShape_VerticalAnchor::TestCases()));

} // namespace gtest_test

void ExShape::InsertTextPaths()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Insert a WordArt object to display text in a shape that we can re-size and move by using the mouse in Microsoft Word.
    // Provide a "ShapeType" as an argument to set a shape for the WordArt.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = AppendWordArt(doc, u"Hello World! This text is bold, and italic.", u"Arial", 480, 24, System::Drawing::Color::get_White(), System::Drawing::Color::get_Black(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    
    // Apply the "Bold" and "Italic" formatting settings to the text using the respective properties.
    shape->get_TextPath()->set_Bold(true);
    shape->get_TextPath()->set_Italic(true);
    
    // Below are various other text formatting-related properties.
    ASSERT_FALSE(shape->get_TextPath()->get_Underline());
    ASSERT_FALSE(shape->get_TextPath()->get_Shadow());
    ASSERT_FALSE(shape->get_TextPath()->get_StrikeThrough());
    ASSERT_FALSE(shape->get_TextPath()->get_ReverseRows());
    ASSERT_FALSE(shape->get_TextPath()->get_XScale());
    ASSERT_FALSE(shape->get_TextPath()->get_Trim());
    ASSERT_FALSE(shape->get_TextPath()->get_SmallCaps());
    
    ASPOSE_ASSERT_EQ(36.0, shape->get_TextPath()->get_Size());
    ASSERT_EQ(u"Hello World! This text is bold, and italic.", shape->get_TextPath()->get_Text());
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::TextPlainText, shape->get_ShapeType());
    
    // Use the "On" property to show/hide the text.
    shape = AppendWordArt(doc, u"On set to \"true\"", u"Calibri", 150, 24, System::Drawing::Color::get_Yellow(), System::Drawing::Color::get_Red(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    shape->get_TextPath()->set_On(true);
    
    shape = AppendWordArt(doc, u"On set to \"false\"", u"Calibri", 150, 24, System::Drawing::Color::get_Yellow(), System::Drawing::Color::get_Purple(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    shape->get_TextPath()->set_On(false);
    
    // Use the "Kerning" property to enable/disable kerning spacing between certain characters.
    shape = AppendWordArt(doc, u"Kerning: VAV", u"Times New Roman", 90, 24, System::Drawing::Color::get_Orange(), System::Drawing::Color::get_Red(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    shape->get_TextPath()->set_Kerning(true);
    
    shape = AppendWordArt(doc, u"No kerning: VAV", u"Times New Roman", 100, 24, System::Drawing::Color::get_Orange(), System::Drawing::Color::get_Red(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    shape->get_TextPath()->set_Kerning(false);
    
    // Use the "Spacing" property to set the custom spacing between characters on a scale from 0.0 (none) to 1.0 (default).
    shape = AppendWordArt(doc, u"Spacing set to 0.1", u"Calibri", 120, 24, System::Drawing::Color::get_BlueViolet(), System::Drawing::Color::get_Blue(), Aspose::Words::Drawing::ShapeType::TextCascadeDown);
    shape->get_TextPath()->set_Spacing(0.1);
    
    // Set the "RotateLetters" property to "true" to rotate each character 90 degrees counterclockwise.
    shape = AppendWordArt(doc, u"RotateLetters", u"Calibri", 200, 36, System::Drawing::Color::get_GreenYellow(), System::Drawing::Color::get_Green(), Aspose::Words::Drawing::ShapeType::TextWave);
    shape->get_TextPath()->set_RotateLetters(true);
    
    // Set the "SameLetterHeights" property to "true" to get the x-height of each character to equal the cap height.
    shape = AppendWordArt(doc, u"Same character height for lower and UPPER case", u"Calibri", 300, 24, System::Drawing::Color::get_DeepSkyBlue(), System::Drawing::Color::get_DodgerBlue(), Aspose::Words::Drawing::ShapeType::TextSlantUp);
    shape->get_TextPath()->set_SameLetterHeights(true);
    
    // By default, the text's size will always scale to fit the containing shape's size, overriding the text size setting.
    shape = AppendWordArt(doc, u"FitShape on", u"Calibri", 160, 24, System::Drawing::Color::get_LightBlue(), System::Drawing::Color::get_Blue(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    ASSERT_TRUE(shape->get_TextPath()->get_FitShape());
    shape->get_TextPath()->set_Size(24.0);
    
    // If we set the "FitShape: property to "false", the text will keep the size
    // which the "Size" property specifies regardless of the size of the shape.
    // Use the "TextPathAlignment" property also to align the text to a side of the shape.
    shape = AppendWordArt(doc, u"FitShape off", u"Calibri", 160, 24, System::Drawing::Color::get_LightBlue(), System::Drawing::Color::get_Blue(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    shape->get_TextPath()->set_FitShape(false);
    shape->get_TextPath()->set_Size(24.0);
    shape->get_TextPath()->set_TextPathAlignment(Aspose::Words::Drawing::TextPathAlignment::Right);
    
    doc->Save(get_ArtifactsDir() + u"Shape.InsertTextPaths.docx");
    TestInsertTextPaths(get_ArtifactsDir() + u"Shape.InsertTextPaths.docx");
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExShape, InsertTextPaths)
{
    s_instance->InsertTextPaths();
}

} // namespace gtest_test

void ExShape::ShapeRevision()
{
    //ExStart
    //ExFor:ShapeBase.IsDeleteRevision
    //ExFor:ShapeBase.IsInsertRevision
    //ExSummary:Shows how to work with revision shapes.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    ASSERT_FALSE(doc->get_TrackRevisions());
    
    // Insert an inline shape without tracking revisions, which will make this shape not a revision of any kind.
    auto shape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Cube);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Inline);
    shape->set_Width(100.0);
    shape->set_Height(100.0);
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(shape);
    
    // Start tracking revisions and then insert another shape, which will be a revision.
    doc->StartTrackRevisions(u"John Doe");
    
    shape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Sun);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Inline);
    shape->set_Width(100.0);
    shape->set_Height(100.0);
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(shape);
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Drawing::Shape>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToArray();
    
    ASSERT_EQ(2, shapes->get_Length());
    
    shapes[0]->Remove();
    
    // Since we removed that shape while we were tracking changes,
    // the shape persists in the document and counts as a delete revision.
    // Accepting this revision will remove the shape permanently, and rejecting it will keep it in the document.
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::Cube, shapes[0]->get_ShapeType());
    ASSERT_TRUE(shapes[0]->get_IsDeleteRevision());
    
    // And we inserted another shape while tracking changes, so that shape will count as an insert revision.
    // Accepting this revision will assimilate this shape into the document as a non-revision,
    // and rejecting the revision will remove this shape permanently.
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::Sun, shapes[1]->get_ShapeType());
    ASSERT_TRUE(shapes[1]->get_IsInsertRevision());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, ShapeRevision)
{
    s_instance->ShapeRevision();
}

} // namespace gtest_test

void ExShape::MoveRevisions()
{
    //ExStart
    //ExFor:ShapeBase.IsMoveFromRevision
    //ExFor:ShapeBase.IsMoveToRevision
    //ExSummary:Shows how to identify move revision shapes.
    // A move revision is when we move an element in the document body by cut-and-pasting it in Microsoft Word while
    // tracking changes. If we involve an inline shape in such a text movement, that shape will also be a revision.
    // Copying-and-pasting or moving floating shapes do not create move revisions.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Revision shape.docx");
    
    // Move revisions consist of pairs of "Move from", and "Move to" revisions. We moved in this document in one shape,
    // but until we accept or reject the move revision, there will be two instances of that shape.
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Drawing::Shape>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToArray();
    
    ASSERT_EQ(2, shapes->get_Length());
    
    // This is the "Move to" revision, which is the shape at its arrival destination.
    // If we accept the revision, this "Move to" revision shape will disappear,
    // and the "Move from" revision shape will remain.
    ASSERT_FALSE(shapes[0]->get_IsMoveFromRevision());
    ASSERT_TRUE(shapes[0]->get_IsMoveToRevision());
    
    // This is the "Move from" revision, which is the shape at its original location.
    // If we accept the revision, this "Move from" revision shape will disappear,
    // and the "Move to" revision shape will remain.
    ASSERT_TRUE(shapes[1]->get_IsMoveFromRevision());
    ASSERT_FALSE(shapes[1]->get_IsMoveToRevision());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, MoveRevisions)
{
    s_instance->MoveRevisions();
}

} // namespace gtest_test

void ExShape::AdjustWithEffects()
{
    //ExStart
    //ExFor:ShapeBase.AdjustWithEffects(RectangleF)
    //ExFor:ShapeBase.BoundsWithEffects
    //ExSummary:Shows how to check how a shape's bounds are affected by shape effects.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Shape shadow effect.docx");
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Drawing::Shape>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToArray();
    
    ASSERT_EQ(2, shapes->get_Length());
    
    // The two shapes are identical in terms of dimensions and shape type.
    ASPOSE_ASSERT_EQ(shapes[0]->get_Width(), shapes[1]->get_Width());
    ASPOSE_ASSERT_EQ(shapes[0]->get_Height(), shapes[1]->get_Height());
    ASSERT_EQ(shapes[0]->get_ShapeType(), shapes[1]->get_ShapeType());
    
    // The first shape has no effects, and the second one has a shadow and thick outline.
    // These effects make the size of the second shape's silhouette bigger than that of the first.
    // Even though the rectangle's size shows up when we click on these shapes in Microsoft Word,
    // the visible outer bounds of the second shape are affected by the shadow and outline and thus are bigger.
    // We can use the "AdjustWithEffects" method to see the true size of the shape.
    ASPOSE_ASSERT_EQ(0.0, shapes[0]->get_StrokeWeight());
    ASPOSE_ASSERT_EQ(20.0, shapes[1]->get_StrokeWeight());
    ASSERT_FALSE(shapes[0]->get_ShadowEnabled());
    ASSERT_TRUE(shapes[1]->get_ShadowEnabled());
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = shapes[0];
    
    // Create a RectangleF object, representing a rectangle,
    // which we could potentially use as the coordinates and bounds for a shape.
    System::Drawing::RectangleF rectangleF(200.0f, 200.0f, 1000.0f, 1000.0f);
    
    // Run this method to get the size of the rectangle adjusted for all our shape effects.
    System::Drawing::RectangleF rectangleFOut = shape->AdjustWithEffects(rectangleF);
    
    // Since the shape has no border-changing effects, its boundary dimensions are unaffected.
    ASPOSE_ASSERT_EQ(200, rectangleFOut.get_X());
    ASPOSE_ASSERT_EQ(200, rectangleFOut.get_Y());
    ASPOSE_ASSERT_EQ(1000, rectangleFOut.get_Width());
    ASPOSE_ASSERT_EQ(1000, rectangleFOut.get_Height());
    
    // Verify the final extent of the first shape, in points.
    ASPOSE_ASSERT_EQ(0, shape->get_BoundsWithEffects().get_X());
    ASPOSE_ASSERT_EQ(0, shape->get_BoundsWithEffects().get_Y());
    ASPOSE_ASSERT_EQ(147, shape->get_BoundsWithEffects().get_Width());
    ASPOSE_ASSERT_EQ(147, shape->get_BoundsWithEffects().get_Height());
    
    shape = shapes[1];
    rectangleF = System::Drawing::RectangleF(200.0f, 200.0f, 1000.0f, 1000.0f);
    rectangleFOut = shape->AdjustWithEffects(rectangleF);
    
    // The shape effects have moved the apparent top left corner of the shape slightly.
    ASPOSE_ASSERT_EQ(171.5, rectangleFOut.get_X());
    ASPOSE_ASSERT_EQ(167, rectangleFOut.get_Y());
    
    // The effects have also affected the visible dimensions of the shape.
    ASPOSE_ASSERT_EQ(1045, rectangleFOut.get_Width());
    ASPOSE_ASSERT_EQ(1133.5, rectangleFOut.get_Height());
    
    // The effects have also affected the visible bounds of the shape.
    ASPOSE_ASSERT_EQ(-28.5, shape->get_BoundsWithEffects().get_X());
    ASPOSE_ASSERT_EQ(-33, shape->get_BoundsWithEffects().get_Y());
    ASPOSE_ASSERT_EQ(192, shape->get_BoundsWithEffects().get_Width());
    ASPOSE_ASSERT_EQ(280.5, shape->get_BoundsWithEffects().get_Height());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, AdjustWithEffects)
{
    s_instance->AdjustWithEffects();
}

} // namespace gtest_test

void ExShape::RenderAllShapes()
{
    //ExStart
    //ExFor:ShapeBase.GetShapeRenderer
    //ExFor:NodeRendererBase.Save(Stream, ImageSaveOptions)
    //ExSummary:Shows how to use a shape renderer to export shapes to files in the local file system.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Various shapes.docx");
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Drawing::Shape>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToArray();
    
    ASSERT_EQ(7, shapes->get_Length());
    
    // There are 7 shapes in the document, including one group shape with 2 child shapes.
    // We will render every shape to an image file in the local file system
    // while ignoring the group shapes since they have no appearance.
    // This will produce 6 image files.
    for (auto&& shape : System::IterateOver(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()))
    {
        System::SharedPtr<Aspose::Words::Rendering::ShapeRenderer> renderer = shape->GetShapeRenderer();
        auto options = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
        renderer->Save(get_ArtifactsDir() + System::String::Format(u"Shape.RenderAllShapes.{0}.png", shape->get_Name()), options);
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, RenderAllShapes)
{
    s_instance->RenderAllShapes();
}

} // namespace gtest_test

void ExShape::DocumentHasSmartArtObject()
{
    //ExStart
    //ExFor:Shape.HasSmartArt
    //ExSummary:Shows how to count the number of shapes in a document with SmartArt objects.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"SmartArt.docx");
    
    int32_t numberOfSmartArtShapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_Cast<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Drawing::Shape>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Drawing::Shape> shape)>>([](System::SharedPtr<Aspose::Words::Drawing::Shape> shape) -> bool
    {
        return shape->get_HasSmartArt();
    })));
    
    ASSERT_EQ(2, numberOfSmartArtShapes);
    //ExEnd
    
}

namespace gtest_test
{

TEST_F(ExShape, DocumentHasSmartArtObject)
{
    s_instance->DocumentHasSmartArtObject();
}

} // namespace gtest_test

void ExShape::OfficeMathRenderer()
{
    //ExStart
    //ExFor:NodeRendererBase
    //ExFor:NodeRendererBase.BoundsInPoints
    //ExFor:NodeRendererBase.GetBoundsInPixels(Single, Single)
    //ExFor:NodeRendererBase.GetBoundsInPixels(Single, Single, Single)
    //ExFor:NodeRendererBase.GetOpaqueBoundsInPixels(Single, Single)
    //ExFor:NodeRendererBase.GetOpaqueBoundsInPixels(Single, Single, Single)
    //ExFor:NodeRendererBase.GetSizeInPixels(Single, Single)
    //ExFor:NodeRendererBase.GetSizeInPixels(Single, Single, Single)
    //ExFor:NodeRendererBase.OpaqueBoundsInPoints
    //ExFor:NodeRendererBase.SizeInPoints
    //ExFor:OfficeMathRenderer
    //ExFor:OfficeMathRenderer.#ctor(OfficeMath)
    //ExSummary:Shows how to measure and scale shapes.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto officeMath = System::ExplicitCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    auto renderer = System::MakeObject<Aspose::Words::Rendering::OfficeMathRenderer>(officeMath);
    
    // Verify the size of the image that the OfficeMath object will create when we render it.
    ASSERT_NEAR(122.0f, renderer->get_SizeInPoints().get_Width(), 0.25f);
    ASSERT_NEAR(13.0f, renderer->get_SizeInPoints().get_Height(), 0.15f);
    
    ASSERT_NEAR(122.0f, renderer->get_BoundsInPoints().get_Width(), 0.25f);
    ASSERT_NEAR(13.0f, renderer->get_BoundsInPoints().get_Height(), 0.15f);
    
    // Shapes with transparent parts may contain different values in the "OpaqueBoundsInPoints" properties.
    ASSERT_NEAR(122.0f, renderer->get_OpaqueBoundsInPoints().get_Width(), 0.25f);
    ASSERT_NEAR(14.2f, renderer->get_OpaqueBoundsInPoints().get_Height(), 0.1f);
    
    // Get the shape size in pixels, with linear scaling to a specific DPI.
    System::Drawing::Rectangle bounds = renderer->GetBoundsInPixels(1.0f, 96.0f);
    
    ASSERT_EQ(163, bounds.get_Width());
    ASSERT_EQ(18, bounds.get_Height());
    
    // Get the shape size in pixels, but with a different DPI for the horizontal and vertical dimensions.
    bounds = renderer->GetBoundsInPixels(1.0f, 96.0f, 150.0f);
    ASSERT_EQ(163, bounds.get_Width());
    ASSERT_EQ(27, bounds.get_Height());
    
    // The opaque bounds may vary here also.
    bounds = renderer->GetOpaqueBoundsInPixels(1.0f, 96.0f);
    
    ASSERT_EQ(163, bounds.get_Width());
    ASSERT_EQ(19, bounds.get_Height());
    
    bounds = renderer->GetOpaqueBoundsInPixels(1.0f, 96.0f, 150.0f);
    
    ASSERT_EQ(163, bounds.get_Width());
    ASSERT_EQ(29, bounds.get_Height());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, SkipMono_OfficeMathRenderer)
{
    RecordProperty("category", "SkipMono");
    
    s_instance->OfficeMathRenderer();
}

} // namespace gtest_test

void ExShape::ShapeTypes()
{
    //ExStart
    //ExFor:ShapeType
    //ExSummary:Shows how Aspose.Words identify shapes.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::Heptagon, Aspose::Words::Drawing::RelativeHorizontalPosition::Page, 0, Aspose::Words::Drawing::RelativeVerticalPosition::Page, 0, 0, 0, Aspose::Words::Drawing::WrapType::None);
    
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::Cloud, Aspose::Words::Drawing::RelativeHorizontalPosition::RightMargin, 0, Aspose::Words::Drawing::RelativeVerticalPosition::Page, 0, 0, 0, Aspose::Words::Drawing::WrapType::None);
    
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::MathPlus, Aspose::Words::Drawing::RelativeHorizontalPosition::RightMargin, 0, Aspose::Words::Drawing::RelativeVerticalPosition::Page, 0, 0, 0, Aspose::Words::Drawing::WrapType::None);
    
    // To correct identify shape types you need to work with shapes as DML.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>(Aspose::Words::SaveFormat::Docx);
    // "Strict" or "Transitional" compliance allows to save shape as DML.
    saveOptions->set_Compliance(Aspose::Words::Saving::OoxmlCompliance::Iso29500_2008_Transitional);
    
    doc->Save(get_ArtifactsDir() + u"Shape.ShapeTypes.docx", saveOptions);
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.ShapeTypes.docx");
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Drawing::Shape>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_ToArray();
    
    for (System::SharedPtr<Aspose::Words::Drawing::Shape> shape : shapes)
    {
        std::cout << System::EnumGetName(shape->get_ShapeType()) << std::endl;
    }
    
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, ShapeTypes)
{
    s_instance->ShapeTypes();
}

} // namespace gtest_test

void ExShape::IsDecorative()
{
    //ExStart
    //ExFor:ShapeBase.IsDecorative
    //ExSummary:Shows how to set that the shape is decorative.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Decorative shapes.docx");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->idx_get(0));
    ASSERT_TRUE(shape->get_IsDecorative());
    
    // If "AlternativeText" is not empty, the shape cannot be decorative.
    // That's why our value has changed to 'false'.
    shape->set_AlternativeText(u"Alternative text.");
    ASSERT_FALSE(shape->get_IsDecorative());
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->MoveToDocumentEnd();
    // Create a new shape as decorative.
    shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 100, 100);
    shape->set_IsDecorative(true);
    
    doc->Save(get_ArtifactsDir() + u"Shape.IsDecorative.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, IsDecorative)
{
    s_instance->IsDecorative();
}

} // namespace gtest_test

void ExShape::FillImage()
{
    //ExStart
    //ExFor:Fill.SetImage(String)
    //ExFor:Fill.SetImage(Byte[])
    //ExFor:Fill.SetImage(Stream)
    //ExSummary:Shows how to set shape fill type as image.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // There are several ways of setting image.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 80, 80);
    // 1 -  Using a local system filename:
    shape->get_Fill()->SetImage(get_ImageDir() + u"Logo.jpg");
    doc->Save(get_ArtifactsDir() + u"Shape.FillImage.FileName.docx");
    
    // 2 -  Load a file into a byte array:
    shape->get_Fill()->SetImage(System::IO::File::ReadAllBytes(get_ImageDir() + u"Logo.jpg"));
    doc->Save(get_ArtifactsDir() + u"Shape.FillImage.ByteArray.docx");
    
    // 3 -  From a stream:
    {
        auto stream = System::MakeObject<System::IO::FileStream>(get_ImageDir() + u"Logo.jpg", System::IO::FileMode::Open);
        shape->get_Fill()->SetImage(stream);
    }
    doc->Save(get_ArtifactsDir() + u"Shape.FillImage.Stream.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, FillImage)
{
    s_instance->FillImage();
}

} // namespace gtest_test

void ExShape::ShadowFormat()
{
    //ExStart
    //ExFor:ShadowFormat.Visible
    //ExFor:ShadowFormat.Clear()
    //ExFor:ShadowType
    //ExSummary:Shows how to work with a shadow formatting for the shape.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Shape stroke pattern border.docx");
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->idx_get(0));
    
    if (shape->get_ShadowFormat()->get_Visible() && shape->get_ShadowFormat()->get_Type() == Aspose::Words::Drawing::ShadowType::Shadow2)
    {
        shape->get_ShadowFormat()->set_Type(Aspose::Words::Drawing::ShadowType::Shadow7);
    }
    
    if (shape->get_ShadowFormat()->get_Type() == Aspose::Words::Drawing::ShadowType::ShadowMixed)
    {
        shape->get_ShadowFormat()->Clear();
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, ShadowFormat)
{
    s_instance->ShadowFormat();
}

} // namespace gtest_test

void ExShape::NoTextRotation()
{
    //ExStart
    //ExFor:TextBox.NoTextRotation
    //ExSummary:Shows how to disable text rotation when the shape is rotate.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Ellipse, 20, 20);
    shape->get_TextBox()->set_NoTextRotation(true);
    
    doc->Save(get_ArtifactsDir() + u"Shape.NoTextRotation.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.NoTextRotation.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->idx_get(0));
    
    ASPOSE_ASSERT_EQ(true, shape->get_TextBox()->get_NoTextRotation());
}

namespace gtest_test
{

TEST_F(ExShape, NoTextRotation)
{
    s_instance->NoTextRotation();
}

} // namespace gtest_test

void ExShape::RelativeSizeAndPosition()
{
    //ExStart
    //ExFor:ShapeBase.RelativeHorizontalSize
    //ExFor:ShapeBase.RelativeVerticalSize
    //ExFor:ShapeBase.WidthRelative
    //ExFor:ShapeBase.HeightRelative
    //ExFor:ShapeBase.TopRelative
    //ExFor:ShapeBase.LeftRelative
    //ExFor:RelativeHorizontalSize
    //ExFor:RelativeVerticalSize
    //ExSummary:Shows how to set relative size and position.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Adding a simple shape with absolute size and position.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 100, 40);
    // Set WrapType to WrapType.None since Inline shapes are automatically converted to absolute units.
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    
    // Checking and setting the relative horizontal size.
    if (shape->get_RelativeHorizontalSize() == Aspose::Words::Drawing::RelativeHorizontalSize::Default)
    {
        // Setting the horizontal size binding to Margin.
        shape->set_RelativeHorizontalSize(Aspose::Words::Drawing::RelativeHorizontalSize::Margin);
        // Setting the width to 50% of Margin width.
        shape->set_WidthRelative(50.0f);
    }
    
    // Checking and setting the relative vertical size.
    if (shape->get_RelativeVerticalSize() == Aspose::Words::Drawing::RelativeVerticalSize::Default)
    {
        // Setting the vertical size binding to Margin.
        shape->set_RelativeVerticalSize(Aspose::Words::Drawing::RelativeVerticalSize::Margin);
        // Setting the heigh to 30% of Margin height.
        shape->set_HeightRelative(30.0f);
    }
    
    // Checking and setting the relative vertical position.
    if (shape->get_RelativeVerticalPosition() == Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph)
    {
        // etting the position binding to TopMargin.
        shape->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::TopMargin);
        // Setting relative Top to 30% of TopMargin position.
        shape->set_TopRelative(30.0f);
    }
    
    // Checking and setting the relative horizontal position.
    if (shape->get_RelativeHorizontalPosition() == Aspose::Words::Drawing::RelativeHorizontalPosition::Default)
    {
        // Setting the position binding to RightMargin.
        shape->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::RightMargin);
        // The position relative value can be negative.
        shape->set_LeftRelative(-260.0f);
    }
    
    doc->Save(get_ArtifactsDir() + u"Shape.RelativeSizeAndPosition.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, RelativeSizeAndPosition)
{
    s_instance->RelativeSizeAndPosition();
}

} // namespace gtest_test

void ExShape::FillBaseColor()
{
    //ExStart:FillBaseColor
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:Fill.BaseForeColor
    //ExFor:Stroke.BaseForeColor
    //ExSummary:Shows how to get foreground color without modifiers.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 100, 40);
    shape->get_Fill()->set_ForeColor(System::Drawing::Color::get_Red());
    shape->get_Fill()->set_ForeTintAndShade(0.5);
    shape->get_Stroke()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Green());
    shape->get_Stroke()->get_Fill()->set_Transparency(0.5);
    
    ASSERT_EQ(System::Drawing::Color::FromArgb(255, 255, 188, 188).ToArgb(), shape->get_Fill()->get_ForeColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), shape->get_Fill()->get_BaseForeColor().ToArgb());
    
    ASSERT_EQ(System::Drawing::Color::FromArgb(128, 0, 128, 0).ToArgb(), shape->get_Stroke()->get_ForeColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), shape->get_Stroke()->get_BaseForeColor().ToArgb());
    
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), shape->get_Stroke()->get_Fill()->get_ForeColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), shape->get_Stroke()->get_Fill()->get_BaseForeColor().ToArgb());
    //ExEnd:FillBaseColor
}

namespace gtest_test
{

TEST_F(ExShape, FillBaseColor)
{
    s_instance->FillBaseColor();
}

} // namespace gtest_test

void ExShape::FitImageToShape()
{
    //ExStart:FitImageToShape
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:ImageData.FitImageToShape
    //ExSummary:Shows hot to fit the image data to Shape frame.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert an image shape and leave its orientation in its default state.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 300, 450);
    shape->get_ImageData()->SetImage(get_ImageDir() + u"Barcode.png");
    shape->get_ImageData()->FitImageToShape();
    
    doc->Save(get_ArtifactsDir() + u"Shape.FitImageToShape.docx");
    //ExEnd:FitImageToShape
}

namespace gtest_test
{

TEST_F(ExShape, FitImageToShape)
{
    s_instance->FitImageToShape();
}

} // namespace gtest_test

void ExShape::StrokeForeThemeColors()
{
    //ExStart:StrokeForeThemeColors
    //GistId:eeeec1fbf118e95e7df3f346c91ed726
    //ExFor:Stroke.ForeThemeColor
    //ExFor:Stroke.ForeTintAndShade
    //ExSummary:Shows how to set fore theme color and tint and shade.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 40);
    System::SharedPtr<Aspose::Words::Drawing::Stroke> stroke = shape->get_Stroke();
    stroke->set_ForeThemeColor(Aspose::Words::Themes::ThemeColor::Dark1);
    stroke->set_ForeTintAndShade(0.5);
    
    doc->Save(get_ArtifactsDir() + u"Shape.StrokeForeThemeColors.docx");
    //ExEnd:StrokeForeThemeColors
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.StrokeForeThemeColors.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeColor::Dark1, shape->get_Stroke()->get_ForeThemeColor());
    ASPOSE_ASSERT_EQ(0.5, shape->get_Stroke()->get_ForeTintAndShade());
}

namespace gtest_test
{

TEST_F(ExShape, StrokeForeThemeColors)
{
    s_instance->StrokeForeThemeColors();
}

} // namespace gtest_test

void ExShape::StrokeBackThemeColors()
{
    //ExStart:StrokeBackThemeColors
    //GistId:eeeec1fbf118e95e7df3f346c91ed726
    //ExFor:Stroke.BackThemeColor
    //ExFor:Stroke.BackTintAndShade
    //ExSummary:Shows how to set back theme color and tint and shade.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Stroke gradient outline.docx");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    System::SharedPtr<Aspose::Words::Drawing::Stroke> stroke = shape->get_Stroke();
    stroke->set_BackThemeColor(Aspose::Words::Themes::ThemeColor::Dark2);
    stroke->set_BackTintAndShade(0.2);
    
    doc->Save(get_ArtifactsDir() + u"Shape.StrokeBackThemeColors.docx");
    //ExEnd:StrokeBackThemeColors
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.StrokeBackThemeColors.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(Aspose::Words::Themes::ThemeColor::Dark2, shape->get_Stroke()->get_BackThemeColor());
    double precision = 1e-6;
    ASSERT_NEAR(0.2, shape->get_Stroke()->get_BackTintAndShade(), precision);
}

namespace gtest_test
{

TEST_F(ExShape, StrokeBackThemeColors)
{
    s_instance->StrokeBackThemeColors();
}

} // namespace gtest_test

void ExShape::TextBoxOleControl()
{
    //ExStart:TextBoxOleControl
    //GistId:eeeec1fbf118e95e7df3f346c91ed726
    //ExFor:TextBoxControl
    //ExFor:TextBoxControl.Text
    //ExFor:TextBoxControl.Type
    //ExSummary:Shows how to change text of the TextBox OLE control.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Textbox control.docm");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    auto textBoxControl = System::ExplicitCast<Aspose::Words::Drawing::Ole::TextBoxControl>(shape->get_OleFormat()->get_OleControl());
    ASSERT_EQ(u"Aspose.Words test", textBoxControl->get_Text());
    
    textBoxControl->set_Text(u"Updated text");
    ASSERT_EQ(u"Updated text", textBoxControl->get_Text());
    ASSERT_EQ(Aspose::Words::Drawing::Ole::Forms2OleControlType::Textbox, textBoxControl->get_Type());
    //ExEnd:TextBoxOleControl
}

namespace gtest_test
{

TEST_F(ExShape, TextBoxOleControl)
{
    s_instance->TextBoxOleControl();
}

} // namespace gtest_test

void ExShape::Glow()
{
    //ExStart:Glow
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:ShapeBase.Glow
    //ExFor:GlowFormat
    //ExFor:GlowFormat.Color
    //ExFor:GlowFormat.Radius
    //ExFor:GlowFormat.Transparency
    //ExFor:GlowFormat.Remove()
    //ExSummary:Shows how to interact with glow shape effect.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Various shapes.docx");
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    shape->get_Glow()->set_Color(System::Drawing::Color::get_Salmon());
    shape->get_Glow()->set_Radius(30);
    shape->get_Glow()->set_Transparency(0.15);
    
    doc->Save(get_ArtifactsDir() + u"Shape.Glow.docx");
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.Glow.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(System::Drawing::Color::FromArgb(217, 250, 128, 114).ToArgb(), shape->get_Glow()->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(30, shape->get_Glow()->get_Radius());
    ASSERT_NEAR(0.15, shape->get_Glow()->get_Transparency(), 0.01);
    
    shape->get_Glow()->Remove();
    
    ASSERT_EQ(System::Drawing::Color::get_Black().ToArgb(), shape->get_Glow()->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(0, shape->get_Glow()->get_Radius());
    ASPOSE_ASSERT_EQ(0, shape->get_Glow()->get_Transparency());
    //ExEnd:Glow
}

namespace gtest_test
{

TEST_F(ExShape, Glow)
{
    s_instance->Glow();
}

} // namespace gtest_test

void ExShape::Reflection()
{
    //ExStart:Reflection
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:ShapeBase.Reflection
    //ExFor:ReflectionFormat
    //ExFor:ReflectionFormat.Size
    //ExFor:ReflectionFormat.Blur
    //ExFor:ReflectionFormat.Transparency
    //ExFor:ReflectionFormat.Distance
    //ExFor:ReflectionFormat.Remove()
    //ExSummary:Shows how to interact with reflection shape effect.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Various shapes.docx");
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    shape->get_Reflection()->set_Transparency(0.37);
    shape->get_Reflection()->set_Size(0.48);
    shape->get_Reflection()->set_Blur(17.5);
    shape->get_Reflection()->set_Distance(9.2);
    
    doc->Save(get_ArtifactsDir() + u"Shape.Reflection.docx");
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.Reflection.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    System::SharedPtr<Aspose::Words::Drawing::ReflectionFormat> reflectionFormat = shape->get_Reflection();
    
    ASSERT_NEAR(0.37, reflectionFormat->get_Transparency(), 0.01);
    ASSERT_NEAR(0.48, reflectionFormat->get_Size(), 0.01);
    ASSERT_NEAR(17.5, reflectionFormat->get_Blur(), 0.01);
    ASSERT_NEAR(9.2, reflectionFormat->get_Distance(), 0.01);
    
    reflectionFormat->Remove();
    
    ASPOSE_ASSERT_EQ(0, reflectionFormat->get_Transparency());
    ASPOSE_ASSERT_EQ(0, reflectionFormat->get_Size());
    ASPOSE_ASSERT_EQ(0, reflectionFormat->get_Blur());
    ASPOSE_ASSERT_EQ(0, reflectionFormat->get_Distance());
    //ExEnd:Reflection
}

namespace gtest_test
{

TEST_F(ExShape, Reflection)
{
    s_instance->Reflection();
}

} // namespace gtest_test

void ExShape::SoftEdge()
{
    //ExStart:SoftEdge
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:ShapeBase.SoftEdge
    //ExFor:SoftEdgeFormat
    //ExFor:SoftEdgeFormat.Radius
    //ExFor:SoftEdgeFormat.Remove
    //ExSummary:Shows how to work with soft edge formatting.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 200, 200);
    
    // Apply soft edge to the shape.
    shape->get_SoftEdge()->set_Radius(30);
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"Shape.SoftEdge.docx");
    
    // Load document with rectangle shape with soft edge.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.SoftEdge.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    System::SharedPtr<Aspose::Words::Drawing::SoftEdgeFormat> softEdgeFormat = shape->get_SoftEdge();
    
    // Check soft edge radius.
    ASPOSE_ASSERT_EQ(30, softEdgeFormat->get_Radius());
    
    // Remove soft edge from the shape.
    softEdgeFormat->Remove();
    
    // Check radius of the removed soft edge.
    ASPOSE_ASSERT_EQ(0, softEdgeFormat->get_Radius());
    //ExEnd:SoftEdge
}

namespace gtest_test
{

TEST_F(ExShape, SoftEdge)
{
    s_instance->SoftEdge();
}

} // namespace gtest_test

void ExShape::Adjustments()
{
    //ExStart:Adjustments
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:Shape.Adjustments
    //ExFor:AdjustmentCollection
    //ExFor:AdjustmentCollection.Count
    //ExFor:AdjustmentCollection.Item(Int32)
    //ExFor:Adjustment
    //ExFor:Adjustment.Name
    //ExFor:Adjustment.Value
    //ExSummary:Shows how to work with adjustment raw values.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rounded rectangle shape.docx");
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    System::SharedPtr<Aspose::Words::Drawing::AdjustmentCollection> adjustments = shape->get_Adjustments();
    ASSERT_EQ(1, adjustments->get_Count());
    
    System::SharedPtr<Aspose::Words::Drawing::Adjustment> adjustment = adjustments->idx_get(0);
    ASSERT_EQ(u"adj", adjustment->get_Name());
    ASSERT_EQ(16667, adjustment->get_Value());
    
    adjustment->set_Value(30000);
    
    doc->Save(get_ArtifactsDir() + u"Shape.Adjustments.docx");
    //ExEnd:Adjustments
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.Adjustments.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    adjustments = shape->get_Adjustments();
    ASSERT_EQ(1, adjustments->get_Count());
    
    adjustment = adjustments->idx_get(0);
    ASSERT_EQ(u"adj", adjustment->get_Name());
    ASSERT_EQ(30000, adjustment->get_Value());
}

namespace gtest_test
{

TEST_F(ExShape, Adjustments)
{
    s_instance->Adjustments();
}

} // namespace gtest_test

void ExShape::ShadowFormatColor()
{
    //ExStart:ShadowFormatColor
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:ShapeBase.ShadowFormat
    //ExFor:ShadowFormat
    //ExFor:ShadowFormat.Color
    //ExFor:ShadowFormat.Type
    //ExSummary:Shows how to get shadow color.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Shadow color.docx");
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    System::SharedPtr<Aspose::Words::Drawing::ShadowFormat> shadowFormat = shape->get_ShadowFormat();
    
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), shadowFormat->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::Drawing::ShadowType::ShadowMixed, shadowFormat->get_Type());
    //ExEnd:ShadowFormatColor
}

namespace gtest_test
{

TEST_F(ExShape, ShadowFormatColor)
{
    s_instance->ShadowFormatColor();
}

} // namespace gtest_test

void ExShape::SetActiveXProperties()
{
    //ExStart:SetActiveXProperties
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:Forms2OleControl.ForeColor
    //ExFor:Forms2OleControl.BackColor
    //ExFor:Forms2OleControl.Height
    //ExFor:Forms2OleControl.Width
    //ExSummary:Shows how to set properties for ActiveX control.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"ActiveX controls.docx");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    auto oleControl = System::ExplicitCast<Aspose::Words::Drawing::Ole::Forms2OleControl>(shape->get_OleFormat()->get_OleControl());
    oleControl->set_ForeColor(System::Drawing::Color::FromArgb(0x17, 0xE1, 0x35));
    oleControl->set_BackColor(System::Drawing::Color::FromArgb(0x33, 0x97, 0xF4));
    oleControl->set_Height(100.54);
    oleControl->set_Width(201.06);
    //ExEnd:SetActiveXProperties
    
    ASSERT_EQ(System::Drawing::Color::FromArgb(0x17, 0xE1, 0x35).ToArgb(), oleControl->get_ForeColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::FromArgb(0x33, 0x97, 0xF4).ToArgb(), oleControl->get_BackColor().ToArgb());
    ASPOSE_ASSERT_EQ(100.54, oleControl->get_Height());
    ASPOSE_ASSERT_EQ(201.06, oleControl->get_Width());
}

namespace gtest_test
{

TEST_F(ExShape, SetActiveXProperties)
{
    s_instance->SetActiveXProperties();
}

} // namespace gtest_test

void ExShape::SelectRadioControl()
{
    //ExStart:SelectRadioControl
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:OptionButtonControl
    //ExFor:OptionButtonControl.Selected
    //ExFor:OptionButtonControl.Type
    //ExSummary:Shows how to select radio button.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Radio buttons.docx");
    
    auto shape1 = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    auto optionButton1 = System::ExplicitCast<Aspose::Words::Drawing::Ole::OptionButtonControl>(shape1->get_OleFormat()->get_OleControl());
    // Deselect selected first item.
    optionButton1->set_Selected(false);
    
    auto shape2 = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    auto optionButton2 = System::ExplicitCast<Aspose::Words::Drawing::Ole::OptionButtonControl>(shape2->get_OleFormat()->get_OleControl());
    // Select second option button.
    optionButton2->set_Selected(true);
    
    ASSERT_EQ(Aspose::Words::Drawing::Ole::Forms2OleControlType::OptionButton, optionButton1->get_Type());
    ASSERT_EQ(Aspose::Words::Drawing::Ole::Forms2OleControlType::OptionButton, optionButton2->get_Type());
    
    doc->Save(get_ArtifactsDir() + u"Shape.SelectRadioControl.docx");
    //ExEnd:SelectRadioControl
}

namespace gtest_test
{

TEST_F(ExShape, SelectRadioControl)
{
    s_instance->SelectRadioControl();
}

} // namespace gtest_test

void ExShape::CheckedCheckBox()
{
    //ExStart:CheckedCheckBox
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:CheckBoxControl
    //ExFor:CheckBoxControl.Checked
    //ExFor:CheckBoxControl.Type
    //ExFor:Forms2OleControlType
    //ExSummary:Shows how to change state of the CheckBox control.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"ActiveX controls.docx");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    auto checkBoxControl = System::ExplicitCast<Aspose::Words::Drawing::Ole::CheckBoxControl>(shape->get_OleFormat()->get_OleControl());
    checkBoxControl->set_Checked(true);
    
    ASPOSE_ASSERT_EQ(true, checkBoxControl->get_Checked());
    ASSERT_EQ(Aspose::Words::Drawing::Ole::Forms2OleControlType::CheckBox, checkBoxControl->get_Type());
    //ExEnd:CheckedCheckBox
}

namespace gtest_test
{

TEST_F(ExShape, CheckedCheckBox)
{
    s_instance->CheckedCheckBox();
}

} // namespace gtest_test

void ExShape::InsertGroupShape()
{
    //ExStart:InsertGroupShape
    //GistId:e06aa7a168b57907a5598e823a22bf0a
    //ExFor:DocumentBuilder.InsertGroupShape(double, double, double, double, ShapeBase[])
    //ExFor:DocumentBuilder.InsertGroupShape(ShapeBase[])
    //ExSummary:Shows how to insert DML group shape.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape1 = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 200, 250);
    shape1->set_Left(20);
    shape1->set_Top(20);
    shape1->get_Stroke()->set_Color(System::Drawing::Color::get_Red());
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape2 = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Ellipse, 150, 200);
    shape2->set_Left(40);
    shape2->set_Top(50);
    shape2->get_Stroke()->set_Color(System::Drawing::Color::get_Green());
    
    // Dimensions for the new GroupShape node.
    double left = 10;
    double top = 10;
    double width = 200;
    double height = 300;
    // Insert GroupShape node for the specified size which is inserted into the specified position.
    System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape1 = builder->InsertGroupShape(left, top, width, height, System::ExplicitCast<System::Array<System::SharedPtr<Aspose::Words::Drawing::ShapeBase>>>(System::MakeArray<System::SharedPtr<Aspose::Words::Drawing::Shape>>({shape1, shape2})));
    
    // Insert GroupShape node which position and dimension will be calculated automatically.
    auto shape3 = System::ExplicitCast<Aspose::Words::Drawing::Shape>(System::ExplicitCast<Aspose::Words::Node>(shape1)->Clone(true));
    System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape2 = builder->InsertGroupShape(System::MakeArray<System::SharedPtr<Aspose::Words::Drawing::ShapeBase>>({shape3}));
    
    doc->Save(get_ArtifactsDir() + u"Shape.InsertGroupShape.docx");
    //ExEnd:InsertGroupShape
}

namespace gtest_test
{

TEST_F(ExShape, InsertGroupShape)
{
    s_instance->InsertGroupShape();
}

} // namespace gtest_test

void ExShape::CombineGroupShape()
{
    //ExStart:CombineGroupShape
    //GistId:bb594993b5fe48692541e16f4d354ac2
    //ExFor:DocumentBuilder.InsertGroupShape(ShapeBase[])
    //ExSummary:Shows how to combine group shape with the shape.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape1 = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 200, 250);
    shape1->set_Left(20);
    shape1->set_Top(20);
    shape1->get_Stroke()->set_Color(System::Drawing::Color::get_Red());
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape2 = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Ellipse, 150, 200);
    shape2->set_Left(40);
    shape2->set_Top(50);
    shape2->get_Stroke()->set_Color(System::Drawing::Color::get_Green());
    
    // Combine shapes into a GroupShape node which is inserted into the specified position.
    System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape1 = builder->InsertGroupShape(System::MakeArray<System::SharedPtr<Aspose::Words::Drawing::ShapeBase>>({shape1, shape2}));
    
    // Combine Shape and GroupShape nodes.
    auto shape3 = System::ExplicitCast<Aspose::Words::Drawing::Shape>(System::ExplicitCast<Aspose::Words::Node>(shape1)->Clone(true));
    System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape2 = builder->InsertGroupShape(System::MakeArray<System::SharedPtr<Aspose::Words::Drawing::ShapeBase>>({groupShape1, shape3}));
    
    doc->Save(get_ArtifactsDir() + u"Shape.CombineGroupShape.docx");
    //ExEnd:CombineGroupShape
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Shape.CombineGroupShape.docx");
    
    System::SharedPtr<Aspose::Words::NodeCollection> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);
    for (auto&& shape : System::IterateOver<Aspose::Words::Drawing::Shape>(shapes))
    {
        ASPOSE_ASSERT_NE(0, shape->get_Width());
        ASPOSE_ASSERT_NE(0, shape->get_Height());
    }
}

namespace gtest_test
{

TEST_F(ExShape, CombineGroupShape)
{
    s_instance->CombineGroupShape();
}

} // namespace gtest_test

void ExShape::InsertCommandButton()
{
    //ExStart:InsertCommandButton
    //GistId:bb594993b5fe48692541e16f4d354ac2
    //ExFor:CommandButtonControl
    //ExFor:CommandButtonControl.#ctor
    //ExFor:CommandButtonControl.Type
    //ExFor:DocumentBuilder.InsertForms2OleControl(Forms2OleControl)
    //ExSummary:Shows how to insert ActiveX control.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    auto button1 = System::MakeObject<Aspose::Words::Drawing::Ole::CommandButtonControl>();
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertForms2OleControl(button1);
    ASSERT_EQ(Aspose::Words::Drawing::Ole::Forms2OleControlType::CommandButton, button1->get_Type());
    //ExEnd:InsertCommandButton
}

namespace gtest_test
{

TEST_F(ExShape, InsertCommandButton)
{
    s_instance->InsertCommandButton();
}

} // namespace gtest_test

void ExShape::Hidden()
{
    //ExStart:Hidden
    //GistId:bb594993b5fe48692541e16f4d354ac2
    //ExFor:ShapeBase.Hidden
    //ExSummary:Shows how to hide the shape.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Shadow color.docx");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    if (!shape->get_Hidden())
    {
        shape->set_Hidden(true);
    }
    
    doc->Save(get_ArtifactsDir() + u"Shape.Hidden.docx");
    //ExEnd:Hidden
}

namespace gtest_test
{

TEST_F(ExShape, Hidden)
{
    s_instance->Hidden();
}

} // namespace gtest_test

void ExShape::CommandButtonCaption()
{
    //ExStart:CommandButtonCaption
    //GistId:366eb64fd56dec3c2eaa40410e594182
    //ExFor:Forms2OleControl.Caption
    //ExSummary:Shows how to set caption for ActiveX control.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    auto button1 = System::MakeObject<Aspose::Words::Drawing::Ole::CommandButtonControl>();
    button1->set_Caption(u"Button caption");
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertForms2OleControl(button1);
    ASSERT_EQ(u"Button caption", button1->get_Caption());
    //ExEnd:CommandButtonCaption
}

namespace gtest_test
{

TEST_F(ExShape, CommandButtonCaption)
{
    s_instance->CommandButtonCaption();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
