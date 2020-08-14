// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExShape.h"

#include <testing/test_predicates.h>
#include <system/type_info.h>
#include <system/text/string_builder.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
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
#include <system/details/dispose_guard.h>
#include <system/console.h>
#include <system/collections/list.h>
#include <system/collections/ienumerable.h>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/size_f.h>
#include <drawing/size.h>
#include <drawing/rectangle_f.h>
#include <drawing/rectangle.h>
#include <drawing/point_f.h>
#include <drawing/point.h>
#include <drawing/image.h>
#include <drawing/color.h>
#include <Aspose.Words.Cpp/Rendering/ShapeRenderer.h>
#include <Aspose.Words.Cpp/Rendering/OfficeMathRenderer.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/HeightRule.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/StoryType.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
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
#include <Aspose.Words.Cpp/Model/Math/MathObjectType.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapSide.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextPathAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextPath.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBoxWrapMode.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBoxAnchor.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBox.h>
#include <Aspose.Words.Cpp/Model/Drawing/Stroke.h>
#include <Aspose.Words.Cpp/Model/Drawing/SignatureLine.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeMarkupLanguage.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeLineStyle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OlePackage.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OleFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OleControl.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/Forms2OleControlType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/Forms2OleControlCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/Forms2OleControl.h>
#include <Aspose.Words.Cpp/Model/Drawing/LayoutFlow.h>
#include <Aspose.Words.Cpp/Model/Drawing/JoinStyle.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Model/Drawing/FlipOrientation.h>
#include <Aspose.Words.Cpp/Model/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Model/Drawing/EndCap.h>
#include <Aspose.Words.Cpp/Model/Drawing/DashStyle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartTitle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/SignatureLineOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "TestUtil.h"
#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Ole;
using namespace Aspose::Words::Math;
using namespace Aspose::Words::Rendering;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1022517304u, ::ApiExamples::ExShape::ShapeVisitor, ThisTypeBaseTypesInfo);

ExShape::ShapeVisitor::ShapeVisitor() : mShapesVisited(0), mTextIndentLevel(0)
{
    mShapesVisited = 0;
    mTextIndentLevel = 0;
    mStringBuilder = MakeObject<System::Text::StringBuilder>();
}

void ExShape::ShapeVisitor::AppendLine(String text)
{
    for (int i = 0; i < mTextIndentLevel; i++)
    {
        mStringBuilder->Append(u'\t');
    }

    mStringBuilder->AppendLine(text);
}

String ExShape::ShapeVisitor::GetText()
{
    return String::Format(u"Shapes visited: {0}\n{1}",mShapesVisited,mStringBuilder);
}

Aspose::Words::VisitorAction ExShape::ShapeVisitor::VisitShapeStart(SharedPtr<Aspose::Words::Drawing::Shape> shape)
{
    AppendLine(String::Format(u"Shape found: {0}",shape->get_ShapeType()));

    mTextIndentLevel++;

    if (shape->get_HasChart())
    {
        AppendLine(String::Format(u"Has chart: {0}",shape->get_Chart()->get_Title()->get_Text()));
    }

    AppendLine(String::Format(u"Extrusion enabled: {0}",shape->get_ExtrusionEnabled()));
    AppendLine(String::Format(u"Shadow enabled: {0}",shape->get_ShadowEnabled()));
    AppendLine(String::Format(u"StoryType: {0}",shape->get_StoryType()));

    if (shape->get_Stroked())
    {
        ASPOSE_EXPECT_EQ(shape->get_Stroke()->get_Color(), shape->get_StrokeColor());
        AppendLine(String::Format(u"Stroke colors: {0}, {1}",shape->get_Stroke()->get_Color(),shape->get_Stroke()->get_Color2()));
        AppendLine(String::Format(u"Stroke weight: {0}",shape->get_StrokeWeight()));

    }

    if (shape->get_Filled())
    {
        AppendLine(String::Format(u"Filled: {0}",shape->get_FillColor()));
    }

    if (shape->get_OleFormat() != nullptr)
    {
        AppendLine(String::Format(u"Ole found of type: {0}",shape->get_OleFormat()->get_ProgId()));
    }

    if (shape->get_SignatureLine() != nullptr)
    {
        AppendLine(String::Format(u"Found signature line for: {0}, {1}",shape->get_SignatureLine()->get_Signer(),shape->get_SignatureLine()->get_SignerTitle()));
    }

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExShape::ShapeVisitor::VisitShapeEnd(SharedPtr<Aspose::Words::Drawing::Shape> shape)
{
    mTextIndentLevel--;
    mShapesVisited++;
    AppendLine(String::Format(u"End of {0}",shape->get_ShapeType()));

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExShape::ShapeVisitor::VisitGroupShapeStart(SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape)
{
    AppendLine(String::Format(u"Shape group found: {0}",groupShape->get_ShapeType()));
    mTextIndentLevel++;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExShape::ShapeVisitor::VisitGroupShapeEnd(SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape)
{
    mTextIndentLevel--;
    AppendLine(String::Format(u"End of {0}",groupShape->get_ShapeType()));

    return Aspose::Words::VisitorAction::Continue;
}

System::Object::shared_members_type ApiExamples::ExShape::ShapeVisitor::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExShape::ShapeVisitor::mStringBuilder", this->mStringBuilder);

    return result;
}

SharedPtr<Aspose::Words::Drawing::Shape> ExShape::AppendWordArt(SharedPtr<Aspose::Words::Document> doc, String text, String textFontFamily, double shapeWidth, double shapeHeight, System::Drawing::Color wordArtFill, System::Drawing::Color line, Aspose::Words::Drawing::ShapeType wordArtShapeType)
{
    // Insert a new paragraph
    auto para = System::DynamicCast<Aspose::Words::Paragraph>(doc->get_FirstSection()->get_Body()->AppendChild(MakeObject<Paragraph>(doc)));

    // Create an inline Shape, which will serve as a container for our WordArt, and append it to the paragraph
    // The shape can only be a valid WordArt shape if the ShapeType assigned here is a WordArt-designated ShapeType
    // These types will have "WordArt object" in the description and their enumerator names will start with "Text..."
    auto shape = MakeObject<Shape>(doc, wordArtShapeType);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Inline);
    para->AppendChild(shape);

    // Set the shape's width and height
    shape->set_Width(shapeWidth);
    shape->set_Height(shapeHeight);

    // These color settings will apply to the letters of the displayed WordArt text
    shape->set_FillColor(wordArtFill);
    shape->set_StrokeColor(line);

    // The WordArt object is accessed here, and we will set the text and font like this
    shape->get_TextPath()->set_Text(text);
    shape->get_TextPath()->set_FontFamily(textFontFamily);

    return shape;
}

void ExShape::TestInsertTextPaths(String filename)
{
    auto doc = MakeObject<Document>(filename);
    SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape> >()->LINQ_ToList();

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, String::Empty, 240, 24, 0.0, 0.0, shapes->idx_get(0));
    ASSERT_TRUE(shapes->idx_get(0)->get_TextPath()->get_Bold());
    ASSERT_TRUE(shapes->idx_get(0)->get_TextPath()->get_Italic());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, String::Empty, 150, 24, 0.0, 0.0, shapes->idx_get(1));
    ASSERT_TRUE(shapes->idx_get(1)->get_TextPath()->get_On());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, String::Empty, 150, 24, 0.0, 0.0, shapes->idx_get(2));
    ASSERT_FALSE(shapes->idx_get(2)->get_TextPath()->get_On());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, String::Empty, 90, 24, 0.0, 0.0, shapes->idx_get(3));
    ASSERT_TRUE(shapes->idx_get(3)->get_TextPath()->get_Kerning());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, String::Empty, 100, 24, 0.0, 0.0, shapes->idx_get(4));
    ASSERT_FALSE(shapes->idx_get(4)->get_TextPath()->get_Kerning());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextCascadeDown, String::Empty, 120, 24, 0.0, 0.0, shapes->idx_get(5));
    ASSERT_NEAR(0.1, shapes->idx_get(5)->get_TextPath()->get_Spacing(), 0.01);

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextWave, String::Empty, 200, 36, 0.0, 0.0, shapes->idx_get(6));
    ASSERT_TRUE(shapes->idx_get(6)->get_TextPath()->get_RotateLetters());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextSlantUp, String::Empty, 300, 24, 0.0, 0.0, shapes->idx_get(7));
    ASSERT_TRUE(shapes->idx_get(7)->get_TextPath()->get_SameLetterHeights());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, String::Empty, 160, 24, 0.0, 0.0, shapes->idx_get(8));
    ASSERT_TRUE(shapes->idx_get(8)->get_TextPath()->get_FitShape());
    ASPOSE_ASSERT_EQ(24.0, shapes->idx_get(8)->get_TextPath()->get_Size());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, String::Empty, 160, 24, 0.0, 0.0, shapes->idx_get(9));
    ASSERT_FALSE(shapes->idx_get(9)->get_TextPath()->get_FitShape());
    ASPOSE_ASSERT_EQ(24.0, shapes->idx_get(9)->get_TextPath()->get_Size());
    ASSERT_EQ(Aspose::Words::Drawing::TextPathAlignment::Right, shapes->idx_get(9)->get_TextPath()->get_TextPathAlignment());
}

namespace gtest_test
{

class ExShape : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExShape> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExShape>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExShape> ExShape::s_instance;

} // namespace gtest_test

void ExShape::Insert()
{
    //ExStart
    //ExFor:ShapeBase.AlternativeText
    //ExFor:ShapeBase.Name
    //ExFor:ShapeBase.Font
    //ExFor:ShapeBase.CanHaveImage
    //ExFor:ShapeBase.ParentParagraph
    //ExFor:ShapeBase.Rotation
    //ExSummary:Shows how to insert shapes.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a cube and set its name
    SharedPtr<Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Cube, 150, 150);
    shape->set_Name(u"MyCube");

    // We can also set the alt text like this
    // This text will be found in Format AutoShape > Alt Text
    shape->set_AlternativeText(u"Alt text for MyCube.");

    // Insert a text box
    shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 300, 50);
    shape->get_Font()->set_Name(u"Times New Roman");

    // Move the builder into the text box and write text
    builder->MoveTo(shape->get_LastParagraph());
    builder->Write(u"Hello world!");

    // Move the builder out of the text box back into the main document
    builder->MoveTo(shape->get_ParentParagraph());

    // Insert a shape with an image
    shape = builder->InsertImage(System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg"));
    ASSERT_TRUE(shape->get_CanHaveImage());
    ASSERT_TRUE(shape->get_HasImage());

    // Rotate the image
    shape->set_Rotation(45.0);

    doc->Save(ArtifactsDir + u"Shape.Insert.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.Insert.docx");
    SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape> >()->LINQ_ToList();

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Cube, u"MyCube", 150.0, 150.0, 0, 0, shapes->idx_get(0));
    ASSERT_EQ(u"Alt text for MyCube.", shapes->idx_get(0)->get_AlternativeText());
    ASSERT_EQ(u"Times New Roman", shapes->idx_get(0)->get_Font()->get_Name());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100004", 300.0, 50.0, 0, 0, shapes->idx_get(1));
    ASSERT_EQ(u"Hello world!", shapes->idx_get(1)->get_LastParagraph()->GetText().Trim());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Image, String::Empty, 300.0, 300.0, 0, 0, shapes->idx_get(2));
    ASSERT_TRUE(shapes->idx_get(2)->get_CanHaveImage());
    ASSERT_TRUE(shapes->idx_get(2)->get_HasImage());
    ASPOSE_ASSERT_EQ(45.0, shapes->idx_get(2)->get_Rotation());
}

namespace gtest_test
{

TEST_F(ExShape, Insert)
{
    s_instance->Insert();
}

} // namespace gtest_test

void ExShape::AspectRatioLockedDefaultValue()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // The best place for the watermark image is in the header or footer so it is shown on every page
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromFile(ImageDir + u"Transparent background logo.png");

    // Insert a floating picture
    SharedPtr<Shape> shape = builder->InsertImage(image);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    shape->set_BehindText(true);

    shape->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::Page);
    shape->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::Page);

    // Calculate image left and top position so it appears in the centre of the page
    shape->set_Left((builder->get_PageSetup()->get_PageWidth() - shape->get_Width()) / 2);
    shape->set_Top((builder->get_PageSetup()->get_PageHeight() - shape->get_Height()) / 2);

    doc = DocumentHelper::SaveOpen(doc);

    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    ASPOSE_ASSERT_EQ(true, shape->get_AspectRatioLocked());
}

namespace gtest_test
{

TEST_F(ExShape, AspectRatioLockedDefaultValue)
{
    s_instance->AspectRatioLockedDefaultValue();
}

} // namespace gtest_test

void ExShape::Coordinates()
{
    //ExStart
    //ExFor:ShapeBase.DistanceBottom
    //ExFor:ShapeBase.DistanceLeft
    //ExFor:ShapeBase.DistanceRight
    //ExFor:ShapeBase.DistanceTop
    //ExSummary:Shows how to set the wrapping distance for text that surrounds a shape.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a rectangle and get the text to wrap tightly around its bounds
    SharedPtr<Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 150, 150);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Tight);

    // Set the minimum distance between the shape and surrounding text
    shape->set_DistanceTop(40.0);
    shape->set_DistanceBottom(40.0);
    shape->set_DistanceLeft(40.0);
    shape->set_DistanceRight(40.0);

    // Move the shape closer to the centre of the page
    shape->set_Top(75.0);
    shape->set_Left(150.0);

    // Rotate the shape
    shape->set_Rotation(60.0);

    // Add text that will wrap around the shape
    builder->get_Font()->set_Size(24.0);
    builder->Write(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ") + u"Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

    doc->Save(ArtifactsDir + u"Shape.Coordinates.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.Coordinates.docx");
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, u"Rectangle 100002", 150.0, 150.0, 75.0, 150.0, shape);
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

void ExShape::InsertGroupShape()
{
    //ExStart
    //ExFor:ShapeBase.AnchorLocked
    //ExFor:ShapeBase.IsTopLevel
    //ExFor:ShapeBase.CoordOrigin
    //ExFor:ShapeBase.CoordSize
    //ExFor:ShapeBase.LocalToParent(PointF)
    //ExSummary:Shows how to create and work with a group of shapes.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    auto group = MakeObject<GroupShape>(doc);

    // Every GroupShape by default is a top level floating shape
    ASSERT_TRUE(group->get_IsGroup());
    ASSERT_TRUE(group->get_IsTopLevel());
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::None, group->get_WrapType());

    // Top level shapes can have this property changed
    group->set_AnchorLocked(true);

    // Set the XY coordinates of the shape group and the size of its containing block, as it appears on the page
    group->set_Bounds(System::Drawing::RectangleF(100.0f, 50.0f, 200.0f, 100.0f));

    // Set the scale of the inner coordinates of the shape group
    // These values mean that the bottom right corner of the 200x100 outer block we set before
    // will be at x = 2000 and y = 1000, or 2000 units from the left and 1000 units from the top
    group->set_CoordSize(System::Drawing::Size(2000, 1000));

    // The coordinate origin of a shape group is x = 0, y = 0 by default, which is the top left corner
    // If we insert a child shape and set its distance from the left to 2000 and the distance from the top to 1000,
    // its origin will be at the bottom right corner of the shape group
    // We can offset the coordinate origin by setting the CoordOrigin attribute
    // In this instance, we move the origin to the centre of the shape group
    group->set_CoordOrigin(System::Drawing::Point(-1000, -500));

    // Populate the shape group with child shapes
    // First, insert a rectangle
    auto subShape = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Rectangle);
    subShape->set_Width(500);
    subShape->set_Height(700);

    // Place its top left corner at the parent group's coordinate origin, which is currently at its centre
    subShape->set_Left(0);
    subShape->set_Top(0);

    // Add the rectangle to the group
    group->AppendChild(subShape);

    // Insert a triangle
    subShape = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Triangle);
    subShape->set_Width(400);
    subShape->set_Height(400);

    // Place its origin at the bottom right corner of the group
    subShape->set_Left(1000);
    subShape->set_Top(500);

    // The offset between this child shape and parent group can be seen here
    ASPOSE_ASSERT_EQ(System::Drawing::PointF(1000.0f, 500.0f), subShape->LocalToParent(System::Drawing::PointF(0.0f, 0.0f)));

    // Add the triangle to the group
    group->AppendChild(subShape);

    // Child shapes of a group shape are not top level
    ASSERT_FALSE(subShape->get_IsTopLevel());

    // Finally, insert the group into the document and save
    builder->InsertNode(group);
    doc->Save(ArtifactsDir + u"Shape.InsertGroupShape.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.InsertGroupShape.docx");
    group = System::DynamicCast<Aspose::Words::Drawing::GroupShape>(doc->GetChild(Aspose::Words::NodeType::GroupShape, 0, true));

    ASSERT_TRUE(group->get_AnchorLocked());
    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(100.0f, 50.0f, 200.0f, 100.0f), group->get_Bounds());
    ASPOSE_ASSERT_EQ(System::Drawing::Size(2000, 1000), group->get_CoordSize());
    ASPOSE_ASSERT_EQ(System::Drawing::Point(-1000, -500), group->get_CoordOrigin());

    subShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(group->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, String::Empty, 500.0, 700.0, 0.0, 0.0, subShape);

    subShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(group->GetChild(Aspose::Words::NodeType::Shape, 1, true));

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Triangle, String::Empty, 400.0, 400.0, 500.0, 1000.0, subShape);
    ASPOSE_ASSERT_EQ(System::Drawing::PointF(1000.0f, 500.0f), subShape->LocalToParent(System::Drawing::PointF(0.0f, 0.0f)));
}

namespace gtest_test
{

TEST_F(ExShape, InsertGroupShape)
{
    s_instance->InsertGroupShape();
}

} // namespace gtest_test

void ExShape::DeleteAllShapes()
{
    //ExStart
    //ExFor:Shape
    //ExSummary:Shows how to delete all shapes from a document.
    // Here we get all shapes from the document node, but you can do this for any smaller
    // node too, for example delete shapes from a single section or a paragraph
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert 2 shapes
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 400, 200);
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::Star, 300, 300);

    // Insert a GroupShape with an inner shape
    auto group = MakeObject<GroupShape>(doc);
    group->set_Bounds(System::Drawing::RectangleF(100.0f, 50.0f, 200.0f, 100.0f));
    group->set_CoordOrigin(System::Drawing::Point(-1000, -500));

    auto subShape = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Cube);
    subShape->set_Width(500);
    subShape->set_Height(700);
    subShape->set_Left(0);
    subShape->set_Top(0);
    group->AppendChild(subShape);
    builder->InsertNode(group);

    ASSERT_EQ(3, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::GroupShape, true)->get_Count());

    // Delete all Shape nodes
    SharedPtr<NodeCollection> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);
    shapes->Clear();

    // The GroupShape node is still present even though there are no sub Shapes
    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::GroupShape, true)->get_Count());
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());

    // GroupShapes also have to be deleted manually
    SharedPtr<NodeCollection> groupShapes = doc->GetChildNodes(Aspose::Words::NodeType::GroupShape, true);
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

void ExShape::CheckShapeInline()
{
    //ExStart
    //ExFor:ShapeBase.IsInline
    //ExSummary:Shows how to test if a shape in the document is inline or floating.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    for (auto shape : System::IterateOver(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape> >()))
    {
        System::Console::WriteLine(shape->get_IsInline() ? String(u"Shape is inline.") : String(u"Shape is floating."));
    }
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);

    ASSERT_FALSE((System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_IsInline());
}

namespace gtest_test
{

TEST_F(ExShape, CheckShapeInline)
{
    s_instance->CheckShapeInline();
}

} // namespace gtest_test

void ExShape::LineFlipOrientation()
{
    //ExStart
    //ExFor:ShapeBase.Bounds
    //ExFor:ShapeBase.BoundsInPoints
    //ExFor:ShapeBase.FlipOrientation
    //ExFor:FlipOrientation
    //ExSummary:Shows how to create line shapes and set specific location and size.
    auto doc = MakeObject<Document>();

    // The lines will cross the whole page
    float pageWidth = (float)doc->get_FirstSection()->get_PageSetup()->get_PageWidth();
    float pageHeight = (float)doc->get_FirstSection()->get_PageSetup()->get_PageHeight();

    // This line goes from top left to bottom right by default
    auto lineA = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Line);
    lineA->set_Bounds(System::Drawing::RectangleF(0.0f, 0.0f, pageWidth, pageHeight));
    lineA->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::Page);
    lineA->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::Page);

    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 0.0f, pageWidth, pageHeight), lineA->get_BoundsInPoints());

    // This line goes from bottom left to top right because we flipped it
    auto lineB = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Line);
    lineB->set_Bounds(System::Drawing::RectangleF(0.0f, 0.0f, pageWidth, pageHeight));
    lineB->set_FlipOrientation(Aspose::Words::Drawing::FlipOrientation::Horizontal);
    lineB->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::Page);
    lineB->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::Page);

    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 0.0f, pageWidth, pageHeight), lineB->get_BoundsInPoints());

    // Add lines to the document
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(lineA);
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(lineB);

    doc->Save(ArtifactsDir + u"Shape.LineFlipOrientation.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.LineFlipOrientation.docx");
    lineA = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 0.0f, pageWidth, pageHeight), lineA->get_BoundsInPoints());
    ASSERT_EQ(Aspose::Words::Drawing::FlipOrientation::None, lineA->get_FlipOrientation());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Page, lineA->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Page, lineA->get_RelativeVerticalPosition());

    lineB = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 0.0f, pageWidth, pageHeight), lineB->get_BoundsInPoints());
    ASSERT_EQ(Aspose::Words::Drawing::FlipOrientation::None, lineB->get_FlipOrientation());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Page, lineB->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Page, lineB->get_RelativeVerticalPosition());
}

namespace gtest_test
{

TEST_F(ExShape, LineFlipOrientation)
{
    s_instance->LineFlipOrientation();
}

} // namespace gtest_test

void ExShape::Fill()
{
    //ExStart
    //ExFor:Shape.Fill
    //ExFor:Shape.FillColor
    //ExFor:Fill
    //ExFor:Fill.Opacity
    //ExSummary:Demonstrates how to create shapes with fill.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Writeln();
    builder->Writeln();
    builder->Writeln();
    builder->Write(u"Some text under the shape.");

    // Create a red balloon, semitransparent
    // The shape is floating and its coordinates are (0,0) by default, relative to the current paragraph
    auto shape = MakeObject<Shape>(builder->get_Document(), Aspose::Words::Drawing::ShapeType::Balloon);
    shape->set_FillColor(System::Drawing::Color::get_Red());
    shape->get_Fill()->set_Opacity(0.3);
    shape->set_Width(100);
    shape->set_Height(100);
    shape->set_Top(-100);
    builder->InsertNode(shape);

    doc->Save(ArtifactsDir + u"Shape.Fill.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.Fill.docx");
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Balloon, String::Empty, 100.0, 100.0, -100.0, 0.0, shape);
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), shape->get_FillColor().ToArgb());
    ASSERT_NEAR(0.3, shape->get_Fill()->get_Opacity(), 0.01);
}

namespace gtest_test
{

TEST_F(ExShape, Fill)
{
    s_instance->Fill();
}

} // namespace gtest_test

void ExShape::Title()
{
    //ExStart
    //ExFor:ShapeBase.Title
    //ExSummary:Shows how to get or set title of shape object.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Create test shape
    auto shape = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Cube);
    shape->set_Width(200);
    shape->set_Height(200);
    shape->set_Title(u"My cube");

    builder->InsertNode(shape);
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Cube, String::Empty, 200.0, 200.0, 0.0, 0.0, shape);
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
    //ExFor:CompositeNode.InsertAfter(Node, Node)
    //ExFor:NodeCollection.ToArray
    //ExSummary:Shows how to replace all textboxes with images.
    auto doc = MakeObject<Document>(MyDir + u"Textboxes in drawing canvas.docx");

    // This gets a live collection of all shape nodes in the document
    SharedPtr<NodeCollection> shapeCollection = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);

    // Since we will be adding/removing nodes, it is better to copy all collection
    // into a fixed size array, otherwise iterator will be invalidated
    ArrayPtr<SharedPtr<Node>> shapes = shapeCollection->ToArray();

    for (auto shape : System::IterateOver(shapes->LINQ_OfType<SharedPtr<Shape> >()))
    {
        // Filter out all shapes that we don't need
        if (System::ObjectExt::Equals(shape->get_ShapeType(), Aspose::Words::Drawing::ShapeType::TextBox))
        {
            // Create a new shape that will replace the existing shape
            auto image = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Image);

            // Load the image into the new shape
            image->get_ImageData()->SetImage(ImageDir + u"Windows MetaFile.wmf");

            // Make new shape's position to match the old shape
            image->set_Left(shape->get_Left());
            image->set_Top(shape->get_Top());
            image->set_Width(shape->get_Width());
            image->set_Height(shape->get_Height());
            image->set_RelativeHorizontalPosition(shape->get_RelativeHorizontalPosition());
            image->set_RelativeVerticalPosition(shape->get_RelativeVerticalPosition());
            image->set_HorizontalAlignment(shape->get_HorizontalAlignment());
            image->set_VerticalAlignment(shape->get_VerticalAlignment());
            image->set_WrapType(shape->get_WrapType());
            image->set_WrapSide(shape->get_WrapSide());

            // Insert new shape after the old shape and remove the old shape
            shape->get_ParentNode()->InsertAfter(image, shape);
            shape->Remove();
        }
    }

    doc->Save(ArtifactsDir + u"Shape.ReplaceTextboxesWithImages.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.ReplaceTextboxesWithImages.docx");
    auto outShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

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
    //ExFor:ShapeBase.ZOrder
    //ExFor:Story.FirstParagraph
    //ExFor:Shape.FirstParagraph
    //ExFor:ShapeBase.WrapType
    //ExSummary:Shows how to create a textbox with some text and different formatting options in a new document.
    auto doc = MakeObject<Document>();

    // Create a new shape of type TextBox
    auto textBox = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::TextBox);

    // Set some settings of the textbox itself
    // Set the wrap of the textbox to inline
    textBox->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    // Set the horizontal and vertical alignment of the text inside the shape
    textBox->set_HorizontalAlignment(Aspose::Words::Drawing::HorizontalAlignment::Center);
    textBox->set_VerticalAlignment(Aspose::Words::Drawing::VerticalAlignment::Top);

    // Set the textbox height and width
    textBox->set_Height(50);
    textBox->set_Width(200);

    // Set the textbox in front of other shapes with a lower ZOrder
    textBox->set_ZOrder(2);

    // Let's create a new paragraph for the textbox manually and align it in the center
    // Make sure we add the new nodes to the textbox as well
    textBox->AppendChild(MakeObject<Paragraph>(doc));
    SharedPtr<Paragraph> para = textBox->get_FirstParagraph();
    para->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);

    // Add some text to the paragraph
    auto run = MakeObject<Run>(doc);
    run->set_Text(u"Hello world!");
    para->AppendChild(run);

    // Append the textbox to the first paragraph in the body
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(textBox);

    doc->Save(ArtifactsDir + u"Shape.CreateTextBox.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.CreateTextBox.docx");
    textBox = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, String::Empty, 200.0, 50.0, 0.0, 0.0, textBox);
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::None, textBox->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::HorizontalAlignment::Center, textBox->get_HorizontalAlignment());
    ASSERT_EQ(Aspose::Words::Drawing::VerticalAlignment::Top, textBox->get_VerticalAlignment());
    ASSERT_EQ(0, textBox->get_ZOrder());
    ASSERT_EQ(u"Hello world!", textBox->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExShape, CreateTextBox)
{
    s_instance->CreateTextBox();
}

} // namespace gtest_test

void ExShape::GetActiveXControlProperties()
{
    //ExStart
    //ExFor:OleControl
    //ExFor:Ole.OleControl.IsForms2OleControl
    //ExFor:Ole.OleControl.Name
    //ExFor:OleFormat.OleControl
    //ExFor:Forms2OleControl
    //ExFor:Forms2OleControl.Caption
    //ExFor:Forms2OleControl.Value
    //ExFor:Forms2OleControl.Enabled
    //ExFor:Forms2OleControl.Type
    //ExFor:Forms2OleControl.ChildNodes
    //ExSummary:Shows how to get ActiveX control and properties from the document.
    auto doc = MakeObject<Document>(MyDir + u"ActiveX controls.docx");

    // Get ActiveX control from the document
    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    SharedPtr<Aspose::Words::Drawing::Ole::OleControl> oleControl = shape->get_OleFormat()->get_OleControl();

    ASPOSE_ASSERT_EQ(nullptr, oleControl->get_Name());

    // Get ActiveX control properties
    if (oleControl->get_IsForms2OleControl())
    {
        auto checkBox = System::DynamicCast<Aspose::Words::Drawing::Ole::Forms2OleControl>(oleControl);
        ASSERT_EQ(u"Первый", checkBox->get_Caption());
        ASSERT_EQ(u"0", checkBox->get_Value());
        ASPOSE_ASSERT_EQ(true, checkBox->get_Enabled());
        ASSERT_EQ(Aspose::Words::Drawing::Ole::Forms2OleControlType::CheckBox, checkBox->get_Type());
        ASPOSE_ASSERT_EQ(nullptr, checkBox->get_ChildNodes());
    }
    //ExEnd
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
    //ExSummary:Shows how to get access to OLE object raw data.
    // Open a document that contains OLE objects
    auto doc = MakeObject<Document>(MyDir + u"OLE objects.docx");

    for (auto shape : System::IterateOver(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)))
    {
        // Get access to OLE data
        SharedPtr<OleFormat> oleFormat = (System::DynamicCast<Aspose::Words::Drawing::Shape>(shape))->get_OleFormat();
        if (oleFormat != nullptr)
        {
            System::Console::WriteLine(String::Format(u"This is {0} object",(oleFormat->get_IsLink() ? String(u"a linked") : String(u"an embedded"))));
            ArrayPtr<uint8_t> oleRawData = oleFormat->GetRawData();
            ASSERT_EQ(24576, oleRawData->get_Length());
            //ExSkip
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
    auto doc = MakeObject<Document>(MyDir + u"OLE spreadsheet.docm");

    // The first shape will contain an OLE object
    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    // This object is a Microsoft Excel spreadsheet
    SharedPtr<OleFormat> oleFormat = shape->get_OleFormat();
    ASSERT_EQ(u"Excel.Sheet.12", oleFormat->get_ProgId());

    // Our object is neither auto updating nor locked from updates
    ASSERT_FALSE(oleFormat->get_AutoUpdate());
    ASPOSE_ASSERT_EQ(false, oleFormat->get_IsLocked());

    // If we want to extract the OLE object by saving it into our local file system, this property can tell us the relevant file extension
    ASSERT_EQ(u".xlsx", oleFormat->get_SuggestedExtension());

    // We can save it via a stream
    {
        auto fs = MakeObject<System::IO::FileStream>(ArtifactsDir + u"OLE spreadsheet extracted via stream" + oleFormat->get_SuggestedExtension(), System::IO::FileMode::Create);
        oleFormat->Save(fs);
    }

    // We can also save it directly to a file
    oleFormat->Save(ArtifactsDir + u"OLE spreadsheet saved directly" + oleFormat->get_SuggestedExtension());
    //ExEnd

    ASSERT_NEAR(8300, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"OLE spreadsheet extracted via stream.xlsx")->get_Length(), TestUtil::get_FileInfoLengthDelta());
    ASSERT_NEAR(8300, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"OLE spreadsheet saved directly.xlsx")->get_Length(), TestUtil::get_FileInfoLengthDelta());
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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Embed a Microsoft Visio drawing as an OLE object into the document
    builder->InsertOleObject(ImageDir + u"Microsoft Visio drawing.vsd", u"Package", false, false, nullptr);

    // Insert a link to the file in the local file system and display it as an icon
    builder->InsertOleObject(ImageDir + u"Microsoft Visio drawing.vsd", u"Package", true, true, nullptr);

    // Both the OLE objects are stored within shapes
    SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape> >()->LINQ_ToList();
    ASSERT_EQ(2, shapes->get_Count());

    // If the shape is an OLE object, it will have a valid OleFormat property
    // We can use it check if it is linked or displayed as an icon, among other things
    SharedPtr<OleFormat> oleFormat = shapes->idx_get(0)->get_OleFormat();
    ASPOSE_ASSERT_EQ(false, oleFormat->get_IsLink());
    ASPOSE_ASSERT_EQ(false, oleFormat->get_OleIcon());

    oleFormat = shapes->idx_get(1)->get_OleFormat();
    ASPOSE_ASSERT_EQ(true, oleFormat->get_IsLink());
    ASPOSE_ASSERT_EQ(true, oleFormat->get_OleIcon());

    // Get the name or the source file and verify that the whole file is linked
    ASSERT_TRUE(oleFormat->get_SourceFullName().EndsWith(String(u"Images") + System::IO::Path::DirectorySeparatorChar + u"Microsoft Visio drawing.vsd"));
    ASSERT_EQ(u"", oleFormat->get_SourceItem());

    ASSERT_EQ(u"Packager", oleFormat->get_IconCaption());

    doc->Save(ArtifactsDir + u"Shape.OleLinks.docx");

    // We can get a stream with the OLE data entry, if the object has this
    {
        SharedPtr<System::IO::MemoryStream> stream = oleFormat->GetOleEntry(u"\x0001" u"CompObj");
        ArrayPtr<uint8_t> oleEntryBytes = stream->ToArray();
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
    //ExFor:Ole.Forms2OleControlCollection
    //ExFor:Ole.Forms2OleControlCollection.Count
    //ExFor:Ole.Forms2OleControlCollection.Item(Int32)
    //ExSummary:Shows how to access an OLE control embedded in a document and its child controls.
    // Open a document that contains a Microsoft Forms OLE control with child controls
    auto doc = MakeObject<Document>(MyDir + u"OLE ActiveX controls.docm");

    // Get the shape that contains the control
    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASSERT_EQ(u"6e182020-f460-11ce-9bcd-00aa00608e01", System::ObjectExt::ToString(shape->get_OleFormat()->get_Clsid()));

    auto oleControl = System::DynamicCast<Aspose::Words::Drawing::Ole::Forms2OleControl>(shape->get_OleFormat()->get_OleControl());

    // Some controls contain child controls
    SharedPtr<Forms2OleControlCollection> oleControlCollection = oleControl->get_ChildNodes();

    // In this case, the child controls are 3 option buttons
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
    //ExSummary:Shows how to get suggested file name from the object.
    auto doc = MakeObject<Document>(MyDir + u"OLE shape.rtf");

    // Gets the file name suggested for the current embedded object if you want to save it into a file
    auto oleShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->get_FirstSection()->get_Body()->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    String suggestedFileName = oleShape->get_OleFormat()->get_SuggestedFileName();

    ASSERT_EQ(u"CSV.csv", suggestedFileName);
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
    auto doc = MakeObject<Document>(MyDir + u"ActiveX controls.docx");

    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    ASSERT_EQ(shape->get_OleFormat()->get_SuggestedFileName().get_Length(), 0);
}

namespace gtest_test
{

TEST_F(ExShape, ObjectDidNotHaveSuggestedFileName)
{
    s_instance->ObjectDidNotHaveSuggestedFileName();
}

} // namespace gtest_test

void ExShape::ResolutionDefaultValues()
{
    auto imageOptions = MakeObject<ImageSaveOptions>(Aspose::Words::SaveFormat::Jpeg);

    ASPOSE_ASSERT_EQ(96, imageOptions->get_HorizontalResolution());
    ASPOSE_ASSERT_EQ(96, imageOptions->get_VerticalResolution());
}

namespace gtest_test
{

TEST_F(ExShape, ResolutionDefaultValues)
{
    s_instance->ResolutionDefaultValues();
}

} // namespace gtest_test

void ExShape::SaveShapeObjectAsImage()
{
    //ExStart
    //ExFor:OfficeMath.GetMathRenderer
    //ExFor:NodeRendererBase.Save(String, ImageSaveOptions)
    //ExSummary:Shows how to convert specific object into image
    auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

    // Get OfficeMath node from the document and render this as image (you can also do the same with the Shape node)
    auto math = System::DynamicCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    math->GetMathRenderer()->Save(ArtifactsDir + u"Shape.SaveShapeObjectAsImage.png", MakeObject<ImageSaveOptions>(Aspose::Words::SaveFormat::Png));
    //ExEnd

    TestUtil::VerifyImage(159, 18, ArtifactsDir + u"Shape.SaveShapeObjectAsImage.png");
}

namespace gtest_test
{

TEST_F(ExShape, SaveShapeObjectAsImage)
{
    s_instance->SaveShapeObjectAsImage();
}

} // namespace gtest_test

void ExShape::OfficeMathDisplayException()
{
    auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

    auto officeMath = System::DynamicCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    officeMath->set_DisplayType(Aspose::Words::Math::OfficeMathDisplayType::Display);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_0 = [&officeMath]()
    {
        officeMath->set_Justification(Aspose::Words::Math::OfficeMathJustification::Inline);
    };

    ASSERT_THROW(_local_func_0(), System::ArgumentException);
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
    auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

    auto officeMath = System::DynamicCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 6, true));

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
    //ExFor:OfficeMath.EquationXmlEncoding
    //ExFor:OfficeMath.Justification
    //ExFor:OfficeMath.NodeType
    //ExFor:OfficeMath.ParentParagraph
    //ExFor:OfficeMathDisplayType
    //ExFor:OfficeMathJustification
    //ExSummary:Shows how to set office math display formatting.
    auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

    auto officeMath = System::DynamicCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));

    // OfficeMath nodes that are children of other OfficeMath nodes are always inline
    // The node we are working with is a base node, so its location and display type can be changed
    ASSERT_EQ(Aspose::Words::Math::MathObjectType::OMathPara, officeMath->get_MathObjectType());
    ASSERT_EQ(Aspose::Words::NodeType::OfficeMath, officeMath->get_NodeType());
    ASPOSE_ASSERT_EQ(officeMath->get_ParentNode(), officeMath->get_ParentParagraph());

    // Used by OOXML and WML formats
    ASSERT_TRUE(officeMath->get_EquationXmlEncoding() == nullptr);

    // We can change the location and display type of the OfficeMath node
    officeMath->set_DisplayType(Aspose::Words::Math::OfficeMathDisplayType::Display);
    officeMath->set_Justification(Aspose::Words::Math::OfficeMathJustification::Left);

    doc->Save(ArtifactsDir + u"Shape.OfficeMath.docx");
    //ExEnd

    ASSERT_TRUE(DocumentHelper::CompareDocs(ArtifactsDir + u"Shape.OfficeMath.docx", GoldsDir + u"Shape.OfficeMath Gold.docx"));
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
    auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

    auto officeMath = System::DynamicCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    officeMath->set_DisplayType(Aspose::Words::Math::OfficeMathDisplayType::Display);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_1 = [&officeMath]()
    {
        officeMath->set_Justification(Aspose::Words::Math::OfficeMathJustification::Inline);
    };

    ASSERT_THROW(_local_func_1(), System::ArgumentException);
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
    auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

    auto officeMath = System::DynamicCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    officeMath->set_DisplayType(Aspose::Words::Math::OfficeMathDisplayType::Inline);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_2 = [&officeMath]()
    {
        officeMath->set_Justification(Aspose::Words::Math::OfficeMathJustification::Center);
    };

    ASSERT_THROW(_local_func_2(), System::ArgumentException);
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
    auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

    auto officeMath = System::DynamicCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));

    // Always inline
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

void ExShape::WorkWithMathObjectType(int index, Aspose::Words::Math::MathObjectType objectType)
{
    auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

    auto officeMath = System::DynamicCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, index, true));
    ASSERT_EQ(objectType, officeMath->get_MathObjectType());
}

namespace gtest_test
{

using WorkWithMathObjectType_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::WorkWithMathObjectType)>::type;

struct WorkWithMathObjectTypeVP : public ExShape, public ApiExamples::ExShape, public ::testing::WithParamInterface<WorkWithMathObjectType_Args>
{
    static std::vector<WorkWithMathObjectType_Args> TestCases()
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

TEST_P(WorkWithMathObjectTypeVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->WorkWithMathObjectType(get<0>(params), get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExShape, WorkWithMathObjectTypeVP, ::testing::ValuesIn(WorkWithMathObjectTypeVP::TestCases()));

} // namespace gtest_test

void ExShape::AspectRatioLocked(bool isLocked)
{
    //ExStart
    //ExFor:ShapeBase.AspectRatioLocked
    //ExSummary:Shows how to set "AspectRatioLocked" for the shape object.
    auto doc = MakeObject<Document>(MyDir + u"ActiveX controls.docx");

    // Get shape object from the document and set AspectRatioLocked(it is possible to get/set AspectRatioLocked for child shapes (mimic MS Word behavior),
    // but AspectRatioLocked has effect only for top level shapes!)
    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    shape->set_AspectRatioLocked(isLocked);
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASPOSE_ASSERT_EQ(isLocked, shape->get_AspectRatioLocked());
}

namespace gtest_test
{

using AspectRatioLocked_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::AspectRatioLocked)>::type;

struct AspectRatioLockedVP : public ExShape, public ApiExamples::ExShape, public ::testing::WithParamInterface<AspectRatioLocked_Args>
{
    static std::vector<AspectRatioLocked_Args> TestCases()
    {
        return
        {
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

} // namespace gtest_test

void ExShape::MarkupLunguageByDefault()
{
    //ExStart
    //ExFor:ShapeBase.MarkupLanguage
    //ExFor:ShapeBase.SizeInPoints
    //ExSummary:Shows how get markup language for shape object in document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->InsertImage(ImageDir + u"Transparent background logo.png");

    // Loop through all single shapes inside document
    for (auto shape : System::IterateOver(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape> >()))
    {
        ASSERT_EQ(Aspose::Words::Drawing::ShapeMarkupLanguage::Dml, shape->get_MarkupLanguage());
        //ExSkip

        System::Console::WriteLine(String(u"Shape: ") + System::ObjectExt::ToString(shape->get_MarkupLanguage()));
        System::Console::WriteLine(String(u"ShapeSize: ") + shape->get_SizeInPoints());
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, MarkupLunguageByDefault)
{
    s_instance->MarkupLunguageByDefault();
}

} // namespace gtest_test

void ExShape::MarkupLunguageForDifferentMsWordVersions(Aspose::Words::Settings::MsWordVersion msWordVersion, Aspose::Words::Drawing::ShapeMarkupLanguage shapeMarkupLanguage)
{
    auto doc = MakeObject<Document>();
    doc->get_CompatibilityOptions()->OptimizeFor(msWordVersion);

    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->InsertImage(ImageDir + u"Transparent background logo.png");

    // Loop through all single shapes inside document
    for (auto shape : System::IterateOver(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape> >()))
    {
        ASSERT_EQ(shapeMarkupLanguage, shape->get_MarkupLanguage());
    }
}

namespace gtest_test
{

using MarkupLunguageForDifferentMsWordVersions_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExShape::MarkupLunguageForDifferentMsWordVersions)>::type;

struct MarkupLunguageForDifferentMsWordVersionsVP : public ExShape, public ApiExamples::ExShape, public ::testing::WithParamInterface<MarkupLunguageForDifferentMsWordVersions_Args>
{
    static std::vector<MarkupLunguageForDifferentMsWordVersions_Args> TestCases()
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

TEST_P(MarkupLunguageForDifferentMsWordVersionsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MarkupLunguageForDifferentMsWordVersions(get<0>(params), get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExShape, MarkupLunguageForDifferentMsWordVersionsVP, ::testing::ValuesIn(MarkupLunguageForDifferentMsWordVersionsVP::TestCases()));

} // namespace gtest_test

void ExShape::ChangeStrokeProperties()
{
    //ExStart
    //ExFor:Stroke
    //ExFor:Stroke.On
    //ExFor:Stroke.Weight
    //ExFor:Stroke.JoinStyle
    //ExFor:Stroke.LineStyle
    //ExFor:ShapeLineStyle
    //ExSummary:Shows how change stroke properties.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Create a new shape of type Rectangle
    auto rectangle = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Rectangle);

    // Change stroke properties
    SharedPtr<Stroke> stroke = rectangle->get_Stroke();
    stroke->set_On(true);
    stroke->set_Weight(5);
    stroke->set_Color(System::Drawing::Color::get_Red());
    stroke->set_DashStyle(Aspose::Words::Drawing::DashStyle::ShortDashDotDot);
    stroke->set_JoinStyle(Aspose::Words::Drawing::JoinStyle::Miter);
    stroke->set_EndCap(Aspose::Words::Drawing::EndCap::Square);
    stroke->set_LineStyle(Aspose::Words::Drawing::ShapeLineStyle::Triple);

    // Insert shape object
    builder->InsertNode(rectangle);
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    rectangle = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    SharedPtr<Stroke> strokeAfter = rectangle->get_Stroke();

    ASPOSE_ASSERT_EQ(true, strokeAfter->get_On());
    ASPOSE_ASSERT_EQ(5, strokeAfter->get_Weight());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), strokeAfter->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::Drawing::DashStyle::ShortDashDotDot, strokeAfter->get_DashStyle());
    ASSERT_EQ(Aspose::Words::Drawing::JoinStyle::Miter, strokeAfter->get_JoinStyle());
    ASSERT_EQ(Aspose::Words::Drawing::EndCap::Square, strokeAfter->get_EndCap());
    ASSERT_EQ(Aspose::Words::Drawing::ShapeLineStyle::Triple, strokeAfter->get_LineStyle());
}

namespace gtest_test
{

TEST_F(ExShape, ChangeStrokeProperties)
{
    s_instance->ChangeStrokeProperties();
}

} // namespace gtest_test

void ExShape::InsertOleObjectAsHtmlFile()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->InsertOleObject(u"http://www.aspose.com", u"htmlfile", true, false, nullptr);

    doc->Save(ArtifactsDir + u"Shape.InsertOleObjectAsHtmlFile.docx");
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
    //ExSummary:Shows how insert ole object as ole package and set it file name and display name.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    ArrayPtr<uint8_t> zipFileBytes = System::IO::File::ReadAllBytes(DatabaseDir + u"cat001.zip");

    {
        auto stream = MakeObject<System::IO::MemoryStream>(zipFileBytes);
        SharedPtr<Shape> shape = builder->InsertOleObject(stream, u"Package", true, nullptr);

        SharedPtr<OlePackage> setOlePackage = shape->get_OleFormat()->get_OlePackage();
        setOlePackage->set_FileName(u"Cat FileName.zip");
        setOlePackage->set_DisplayName(u"Cat DisplayName.zip");

        doc->Save(ArtifactsDir + u"Shape.InsertOlePackage.docx");
    }
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.InsertOlePackage.docx");

    auto getShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    SharedPtr<OlePackage> getOlePackage = getShape->get_OleFormat()->get_OlePackage();

    ASSERT_EQ(u"Cat FileName.zip", getOlePackage->get_FileName());
    ASSERT_EQ(u"Cat DisplayName.zip", getOlePackage->get_DisplayName());
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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<Shape> oleObject = builder->InsertOleObject(MyDir + u"Spreadsheet.xlsx", false, false, nullptr);
    SharedPtr<Shape> oleObjectAsOlePackage = builder->InsertOleObject(MyDir + u"Spreadsheet.xlsx", u"Excel.Sheet", false, false, nullptr);

    ASPOSE_ASSERT_EQ(nullptr, oleObject->get_OleFormat()->get_OlePackage());
    ASPOSE_ASSERT_EQ(System::ObjectExt::GetType<OlePackage>(), System::ObjectExt::GetType(oleObjectAsOlePackage->get_OleFormat()->get_OlePackage()));
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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 200, 300);
    // Change shape size and rotation
    shape->set_Height(300);
    shape->set_Width(500);
    shape->set_Rotation(30);

    doc->Save(ArtifactsDir + u"Shape.Resize.docx");
}

namespace gtest_test
{

TEST_F(ExShape, Resize)
{
    s_instance->Resize();
}

} // namespace gtest_test

void ExShape::LayoutInTableCell()
{
    //ExStart
    //ExFor:ShapeBase.IsLayoutInCell
    //ExFor:MsWordVersion
    //ExSummary:Shows how to display the shape, inside a table or outside of it.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->StartTable();
    builder->get_RowFormat()->set_Height(100);
    builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Exactly);

    for (int i = 0; i < 31; i++)
    {
        if (i != 0 && i % 7 == 0)
        {
            builder->EndRow();
        }
        builder->InsertCell();
        builder->Write(u"Cell contents");
    }

    builder->EndTable();

    SharedPtr<NodeCollection> runs = doc->GetChildNodes(Aspose::Words::NodeType::Run, true);
    int num = 1;

    for (auto run : System::IterateOver(runs->LINQ_OfType<SharedPtr<Run> >()))
    {
        auto watermark = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::TextPlainText);
        watermark->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::Page);
        watermark->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::Page);
        // False - display the shape outside of table cell, True - display the shape outside of table cell
        watermark->set_IsLayoutInCell(true);

        watermark->set_Width(30);
        watermark->set_Height(30);
        watermark->set_HorizontalAlignment(Aspose::Words::Drawing::HorizontalAlignment::Center);
        watermark->set_VerticalAlignment(Aspose::Words::Drawing::VerticalAlignment::Center);

        watermark->set_Rotation(-40);
        watermark->get_Fill()->set_Color(System::Drawing::Color::get_Gainsboro());
        watermark->set_StrokeColor(System::Drawing::Color::get_Gainsboro());

        watermark->get_TextPath()->set_Text(String::Format(u"{0}",num));
        watermark->get_TextPath()->set_FontFamily(u"Arial");

        watermark->set_Name(String::Format(u"Watermark_{0}",num++));
        // Property will take effect only if the WrapType property is set to something other than WrapType.Inline
        watermark->set_WrapType(Aspose::Words::Drawing::WrapType::None);
        watermark->set_BehindText(true);

        builder->MoveTo(run);
        builder->InsertNode(watermark);
    }

    // Behaviour of MS Word on working with shapes in table cells is changed in the last versions
    // Adding the following line is needed to make the shape displayed in center of a page
    doc->get_CompatibilityOptions()->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2010);

    doc->Save(ArtifactsDir + u"Shape.LayoutInTableCell.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.LayoutInTableCell.docx");
    SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape> >()->LINQ_ToList();

    ASSERT_EQ(31, shapes->get_Count());

    for (auto shape : shapes)
    {
        TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextPlainText, String::Format(u"Watermark_{0}",shapes->IndexOf(shape) + 1), 30.0, 30.0, 0.0, 0.0, shape);
    }
}

namespace gtest_test
{

TEST_F(ExShape, LayoutInTableCell)
{
    s_instance->LayoutInTableCell();
}

} // namespace gtest_test

void ExShape::ShapeInsertion()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertShape(ShapeType, RelativeHorizontalPosition, double, RelativeVerticalPosition, double, double, double, WrapType)
    //ExFor:DocumentBuilder.InsertShape(ShapeType, double, double)
    //ExSummary:Shows how to insert DML shapes into the document using a document builder.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // There are two ways of shape insertion
    // These methods allow inserting DML shape into the document model
    // Document must be saved in the format, which supports DML shapes, otherwise, such nodes will be converted
    // to VML shape, while document saving

    // 1. Free-floating shape insertion
    SharedPtr<Shape> freeFloatingShape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TopCornersRounded, Aspose::Words::Drawing::RelativeHorizontalPosition::Page, 100, Aspose::Words::Drawing::RelativeVerticalPosition::Page, 100, 50, 50, Aspose::Words::Drawing::WrapType::None);
    freeFloatingShape->set_Rotation(30.0);
    // 2. Inline shape insertion
    SharedPtr<Shape> inlineShape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::DiagonalCornersRounded, 50, 50);
    inlineShape->set_Rotation(30.0);

    // If you need to create "NonPrimitive" shapes, like SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
    // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
    // please save the document with "Strict" or "Transitional" compliance which allows saving shape as DML
    auto saveOptions = MakeObject<OoxmlSaveOptions>(Aspose::Words::SaveFormat::Docx);
    saveOptions->set_Compliance(Aspose::Words::Saving::OoxmlCompliance::Iso29500_2008_Transitional);

    doc->Save(ArtifactsDir + u"Shape.ShapeInsertion.docx", saveOptions);
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.ShapeInsertion.docx");
    SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape> >()->LINQ_ToList();

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TopCornersRounded, u"TopCornersRounded 100002", 50.0, 50.0, 100.0, 100.0, shapes->idx_get(0));
    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::DiagonalCornersRounded, u"DiagonalCornersRounded 100004", 50.0, 50.0, 0.0, 0.0, shapes->idx_get(1));
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
    // Open a document that contains shapes
    auto doc = MakeObject<Document>(MyDir + u"Revision shape.docx");
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    //ExSKip

    // Create a ShapeVisitor and get the document to accept it
    auto shapeVisitor = MakeObject<ExShape::ShapeVisitor>();
    doc->Accept(shapeVisitor);

    // Print all the information that the visitor has collected
    System::Console::WriteLine(shapeVisitor->GetText());
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
    //ExFor:SignatureLine.IsSigned
    //ExFor:SignatureLine.IsValid
    //ExFor:SignatureLine.ShowDate
    //ExFor:SignatureLine.Signer
    //ExFor:SignatureLine.SignerTitle
    //ExSummary:Shows how to create a line for a signature and insert it into a document.
    // Create a blank document and its DocumentBuilder
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // The SignatureLineOptions will contain all the data that the signature line will display
    auto options = MakeObject<SignatureLineOptions>();
    options->set_AllowComments(true);
    options->set_DefaultInstructions(true);
    options->set_Email(u"john.doe@management.com");
    options->set_Instructions(u"Please sign here");
    options->set_ShowDate(true);
    options->set_Signer(u"John Doe");
    options->set_SignerTitle(u"Senior Manager");

    // Insert the signature line, applying our SignatureLineOptions
    // We can control where the signature line will appear on the page using a combination of left/top indents and margin-relative positions
    // Since we're placing the signature line at the bottom right of the page, we will need to use negative indents to move it into view
    SharedPtr<Shape> shape = builder->InsertSignatureLine(options, Aspose::Words::Drawing::RelativeHorizontalPosition::RightMargin, -170.0, Aspose::Words::Drawing::RelativeVerticalPosition::BottomMargin, -60.0, Aspose::Words::Drawing::WrapType::None);
    ASSERT_TRUE(shape->get_IsSignatureLine());

    // The SignatureLine object is a member of the shape that contains it
    SharedPtr<Aspose::Words::Drawing::SignatureLine> signatureLine = shape->get_SignatureLine();

    ASSERT_EQ(u"john.doe@management.com", signatureLine->get_Email());
    ASSERT_EQ(u"John Doe", signatureLine->get_Signer());
    ASSERT_EQ(u"Senior Manager", signatureLine->get_SignerTitle());
    ASSERT_EQ(u"Please sign here", signatureLine->get_Instructions());
    ASSERT_TRUE(signatureLine->get_ShowDate());
    ASSERT_TRUE(signatureLine->get_AllowComments());
    ASSERT_TRUE(signatureLine->get_DefaultInstructions());

    // We will be prompted to sign it when we open the document
    ASSERT_FALSE(signatureLine->get_IsSigned());

    // The object may be valid, but the signature itself isn't until it is signed
    ASSERT_FALSE(signatureLine->get_IsValid());

    doc->Save(ArtifactsDir + u"Shape.SignatureLine.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.SignatureLine.docx");
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Image, String::Empty, 192.75, 96.75, -60.0, -170.0, shape);
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

void ExShape::TextBox()
{
    //ExStart
    //ExFor:Shape.TextBox
    //ExFor:Shape.LastParagraph
    //ExFor:TextBox
    //ExFor:TextBox.FitShapeToText
    //ExFor:TextBox.InternalMarginBottom
    //ExFor:TextBox.InternalMarginLeft
    //ExFor:TextBox.InternalMarginRight
    //ExFor:TextBox.InternalMarginTop
    //ExFor:TextBox.LayoutFlow
    //ExFor:TextBox.TextBoxWrapMode
    //ExFor:TextBoxWrapMode
    //ExSummary:Shows how to insert text boxes and arrange their text.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a shape that contains a TextBox
    SharedPtr<Shape> textBoxShape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 150, 100);
    SharedPtr<Aspose::Words::Drawing::TextBox> textBox = textBoxShape->get_TextBox();

    // Move the document builder to inside the TextBox and write text
    builder->MoveTo(textBoxShape->get_LastParagraph());
    builder->Write(u"Vertical text");

    // Text is displayed vertically, written top to bottom
    textBox->set_LayoutFlow(Aspose::Words::Drawing::LayoutFlow::TopToBottomIdeographic);

    // Move the builder out of the shape and back into the main document body
    builder->MoveTo(textBoxShape->get_ParentParagraph());

    // Insert another TextBox
    textBoxShape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 150, 100);
    textBox = textBoxShape->get_TextBox();

    // Apply these values to both these members to get the parent shape to defy the dimensions we set to fit tightly around the TextBox's text
    textBox->set_FitShapeToText(true);
    textBox->set_TextBoxWrapMode(Aspose::Words::Drawing::TextBoxWrapMode::None);

    builder->MoveTo(textBoxShape->get_LastParagraph());
    builder->Write(u"Text fit tightly inside textbox");

    builder->MoveTo(textBoxShape->get_ParentParagraph());

    textBoxShape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 100);
    textBox = textBoxShape->get_TextBox();

    // Set margins for the textbox
    textBox->set_InternalMarginTop(15);
    textBox->set_InternalMarginBottom(15);
    textBox->set_InternalMarginLeft(15);
    textBox->set_InternalMarginRight(15);

    builder->MoveTo(textBoxShape->get_LastParagraph());
    builder->Write(u"Text placed according to textbox margins");

    doc->Save(ArtifactsDir + u"Shape.TextBox.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.TextBox.docx");
    SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape> >()->LINQ_ToList();

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100002", 150.0, 100.0, 0.0, 0.0, shapes->idx_get(0));
    TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::TopToBottomIdeographic, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(0)->get_TextBox());
    ASSERT_EQ(u"Vertical text", shapes->idx_get(0)->GetText().Trim());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100004", 150.0, 100.0, 0.0, 0.0, shapes->idx_get(1));
    TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, true, Aspose::Words::Drawing::TextBoxWrapMode::None, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(1)->get_TextBox());
    ASSERT_EQ(u"Text fit tightly inside textbox", shapes->idx_get(1)->GetText().Trim());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100006", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(2));
    TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 15.0, 15.0, 15.0, 15.0, shapes->idx_get(2)->get_TextBox());
    ASSERT_EQ(u"Text placed according to textbox margins", shapes->idx_get(2)->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExShape, TextBox)
{
    s_instance->TextBox();
}

} // namespace gtest_test

void ExShape::TextBoxShapeType()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set compatibility options to correctly using of VerticalAnchor property
    doc->get_CompatibilityOptions()->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2016);

    SharedPtr<Shape> textBoxShape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 100);
    // Not all formats are compatible with this one
    // For most of incompatible formats AW generated a warnings on save, so use doc.WarningCallback to check it
    textBoxShape->get_TextBox()->set_VerticalAnchor(Aspose::Words::Drawing::TextBoxAnchor::Bottom);

    builder->MoveTo(textBoxShape->get_LastParagraph());
    builder->Write(u"Text placed bottom");

    doc->Save(ArtifactsDir + u"Shape.TextBoxShapeType.docx");
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
    //ExSummary:Shows how to work with textbox forward link
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Create a few textboxes for example
    SharedPtr<Shape> textBoxShape1 = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 100);
    SharedPtr<Aspose::Words::Drawing::TextBox> textBox1 = textBoxShape1->get_TextBox();
    builder->Writeln();

    SharedPtr<Shape> textBoxShape2 = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 100);
    SharedPtr<Aspose::Words::Drawing::TextBox> textBox2 = textBoxShape2->get_TextBox();
    builder->Writeln();

    SharedPtr<Shape> textBoxShape3 = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 100);
    SharedPtr<Aspose::Words::Drawing::TextBox> textBox3 = textBoxShape3->get_TextBox();
    builder->Writeln();

    SharedPtr<Shape> textBoxShape4 = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 100);
    SharedPtr<Aspose::Words::Drawing::TextBox> textBox4 = textBoxShape4->get_TextBox();

    // Create link between textboxes if possible
    if (textBox1->IsValidLinkTarget(textBox2))
    {
        textBox1->set_Next(textBox2);
    }

    if (textBox2->IsValidLinkTarget(textBox3))
    {
        textBox2->set_Next(textBox3);
    }

    // You can only create link on empty textbox
    builder->MoveTo(textBoxShape4->get_LastParagraph());
    builder->Write(u"Vertical text");
    // Thus it's not valid link target
    ASSERT_FALSE(textBox3->IsValidLinkTarget(textBox4));

    if (textBox1->get_Next() != nullptr && textBox1->get_Previous() == nullptr)
    {
        System::Console::WriteLine(u"This TextBox is the head of the sequence");
    }

    if (textBox2->get_Next() != nullptr && textBox2->get_Previous() != nullptr)
    {
        System::Console::WriteLine(u"This TextBox is the middle of the sequence");
    }

    if (textBox3->get_Next() == nullptr && textBox3->get_Previous() != nullptr)
    {
        System::Console::WriteLine(u"This TextBox is the tail of the sequence");

        // Break the forward link between textBox2 and textBox3
        textBox3->get_Previous()->BreakForwardLink();
        // Check that link was break successfully
        ASSERT_TRUE(textBox2->get_Next() == nullptr);
        ASSERT_TRUE(textBox3->get_Previous() == nullptr);
    }

    doc->Save(ArtifactsDir + u"Shape.CreateLinkBetweenTextBoxes.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.CreateLinkBetweenTextBoxes.docx");
    SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape> >()->LINQ_ToList();

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100002", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(0));
    TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(0)->get_TextBox());
    ASSERT_EQ(String::Empty, shapes->idx_get(0)->GetText().Trim());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100004", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(1));
    TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(1)->get_TextBox());
    ASSERT_EQ(String::Empty, shapes->idx_get(1)->GetText().Trim());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::Rectangle, u"TextBox 100006", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(2));
    TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(2)->get_TextBox());
    ASSERT_EQ(String::Empty, shapes->idx_get(2)->GetText().Trim());

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100008", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(3));
    TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(3)->get_TextBox());
    ASSERT_EQ(u"Vertical text", shapes->idx_get(3)->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExShape, CreateLinkBetweenTextBoxes)
{
    s_instance->CreateLinkBetweenTextBoxes();
}

} // namespace gtest_test

void ExShape::GetTextBoxAndChangeTextAnchor()
{
    //ExStart
    //ExFor:TextBoxAnchor
    //ExFor:TextBox.VerticalAnchor
    //ExSummary:Shows how to change text position inside textbox shape.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<Shape> textBox = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 200, 200);
    textBox->get_TextBox()->set_VerticalAnchor(Aspose::Words::Drawing::TextBoxAnchor::Bottom);

    builder->MoveTo(textBox->get_FirstParagraph());
    builder->Write(u"Textbox contents");

    doc->Save(ArtifactsDir + u"Shape.GetTextBoxAndChangeAnchor.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Shape.GetTextBoxAndChangeAnchor.docx");
    textBox = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType::TextBox, u"TextBox 100002", 200.0, 200.0, 0.0, 0.0, textBox);
    TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow::Horizontal, false, Aspose::Words::Drawing::TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, textBox->get_TextBox());
    ASSERT_EQ(u"Textbox contents", textBox->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExShape, GetTextBoxAndChangeTextAnchor)
{
    s_instance->GetTextBoxAndChangeTextAnchor();
}

} // namespace gtest_test

void ExShape::InsertTextPaths()
{
    auto doc = MakeObject<Document>();

    // Insert a WordArt object and capture the shape that contains it in a variable
    SharedPtr<Shape> shape = AppendWordArt(doc, u"Bold & Italic", u"Arial", 240, 24, System::Drawing::Color::get_White(), System::Drawing::Color::get_Black(), Aspose::Words::Drawing::ShapeType::TextPlainText);

    // View and verify various text formatting settings
    shape->get_TextPath()->set_Bold(true);
    shape->get_TextPath()->set_Italic(true);

    ASSERT_FALSE(shape->get_TextPath()->get_Underline());
    ASSERT_FALSE(shape->get_TextPath()->get_Shadow());
    ASSERT_FALSE(shape->get_TextPath()->get_StrikeThrough());
    ASSERT_FALSE(shape->get_TextPath()->get_ReverseRows());
    ASSERT_FALSE(shape->get_TextPath()->get_XScale());
    ASSERT_FALSE(shape->get_TextPath()->get_Trim());
    ASSERT_FALSE(shape->get_TextPath()->get_SmallCaps());

    ASPOSE_ASSERT_EQ(36.0, shape->get_TextPath()->get_Size());
    ASSERT_EQ(u"Bold & Italic", shape->get_TextPath()->get_Text());
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::TextPlainText, shape->get_ShapeType());

    // Toggle whether or not to display text
    shape = AppendWordArt(doc, u"On set to true", u"Calibri", 150, 24, System::Drawing::Color::get_Yellow(), System::Drawing::Color::get_Purple(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    shape->get_TextPath()->set_On(true);

    shape = AppendWordArt(doc, u"On set to false", u"Calibri", 150, 24, System::Drawing::Color::get_Yellow(), System::Drawing::Color::get_Purple(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    shape->get_TextPath()->set_On(false);

    // Apply kerning
    shape = AppendWordArt(doc, u"Kerning: VAV", u"Times New Roman", 90, 24, System::Drawing::Color::get_Orange(), System::Drawing::Color::get_Red(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    shape->get_TextPath()->set_Kerning(true);

    shape = AppendWordArt(doc, u"No kerning: VAV", u"Times New Roman", 100, 24, System::Drawing::Color::get_Orange(), System::Drawing::Color::get_Red(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    shape->get_TextPath()->set_Kerning(false);

    // Apply custom spacing, on a scale from 0.0 (none) to 1.0 (default)
    shape = AppendWordArt(doc, u"Spacing set to 0.1", u"Calibri", 120, 24, System::Drawing::Color::get_BlueViolet(), System::Drawing::Color::get_Blue(), Aspose::Words::Drawing::ShapeType::TextCascadeDown);
    shape->get_TextPath()->set_Spacing(0.1);

    // Rotate letters 90 degrees to the left, text is still laid out horizontally
    shape = AppendWordArt(doc, u"RotateLetters", u"Calibri", 200, 36, System::Drawing::Color::get_GreenYellow(), System::Drawing::Color::get_Green(), Aspose::Words::Drawing::ShapeType::TextWave);
    shape->get_TextPath()->set_RotateLetters(true);

    // Set the x-height to equal the cap height
    shape = AppendWordArt(doc, u"Same character height for lower and UPPER case", u"Calibri", 300, 24, System::Drawing::Color::get_DeepSkyBlue(), System::Drawing::Color::get_DodgerBlue(), Aspose::Words::Drawing::ShapeType::TextSlantUp);
    shape->get_TextPath()->set_SameLetterHeights(true);

    // By default, the size of the text will scale to always fit the size of the containing shape, overriding the text size setting
    shape = AppendWordArt(doc, u"FitShape on", u"Calibri", 160, 24, System::Drawing::Color::get_LightBlue(), System::Drawing::Color::get_Blue(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    ASSERT_TRUE(shape->get_TextPath()->get_FitShape());
    shape->get_TextPath()->set_Size(24.0);

    // If we set FitShape to false, the size of the text will defy the shape bounds and always keep the size value we set below
    // We can also set TextPathAlignment to align the text
    shape = AppendWordArt(doc, u"FitShape off", u"Calibri", 160, 24, System::Drawing::Color::get_LightBlue(), System::Drawing::Color::get_Blue(), Aspose::Words::Drawing::ShapeType::TextPlainText);
    shape->get_TextPath()->set_FitShape(false);
    shape->get_TextPath()->set_Size(24.0);
    shape->get_TextPath()->set_TextPathAlignment(Aspose::Words::Drawing::TextPathAlignment::Right);

    doc->Save(ArtifactsDir + u"Shape.InsertTextPaths.docx");
    TestInsertTextPaths(ArtifactsDir + u"Shape.InsertTextPaths.docx");
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
    // Open a blank document
    auto doc = MakeObject<Document>();

    // Insert an inline shape without tracking revisions
    ASSERT_FALSE(doc->get_TrackRevisions());
    auto shape = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Cube);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Inline);
    shape->set_Width(100.0);
    shape->set_Height(100.0);
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(shape);

    // Start tracking revisions and then insert another shape
    doc->StartTrackRevisions(u"John Doe");

    shape = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Sun);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Inline);
    shape->set_Width(100.0);
    shape->set_Height(100.0);
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(shape);

    // Get the document's shape collection which includes just the two shapes we added
    SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape> >()->LINQ_ToList();
    ASSERT_EQ(2, shapes->get_Count());

    // Remove the first shape
    shapes->idx_get(0)->Remove();

    // Because we removed that shape while changes were being tracked, the shape counts as a delete revision
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::Cube, shapes->idx_get(0)->get_ShapeType());
    ASSERT_TRUE(shapes->idx_get(0)->get_IsDeleteRevision());

    // And we inserted another shape while tracking changes, so that shape will count as an insert revision
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::Sun, shapes->idx_get(1)->get_ShapeType());
    ASSERT_TRUE(shapes->idx_get(1)->get_IsInsertRevision());
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
    // Open a document that contains a move revision
    // A move revision is when we, while changes are tracked, cut(not copy)-and-paste or highlight and drag text from one place to another
    // If inline shapes are caught up in the text movement, they will count as move revisions as well
    // Moving a floating shape will not count as a move revision
    auto doc = MakeObject<Document>(MyDir + u"Revision shape.docx");

    // The document has one shape that was moved, but shape move revisions will have two instances of that shape
    // One will be the shape at its arrival destination and the other will be the shape at its original location
    SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> nc = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape> >()->LINQ_ToList();
    ASSERT_EQ(2, nc->get_Count());

    // This is the move to revision, also the shape at its arrival destination
    ASSERT_FALSE(nc->idx_get(0)->get_IsMoveFromRevision());
    ASSERT_TRUE(nc->idx_get(0)->get_IsMoveToRevision());

    // This is the move from revision, which is the shape at its original location
    ASSERT_TRUE(nc->idx_get(1)->get_IsMoveFromRevision());
    ASSERT_FALSE(nc->idx_get(1)->get_IsMoveToRevision());
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
    // Open a document that contains two shapes and get its shape collection
    auto doc = MakeObject<Document>(MyDir + u"Shape shadow effect.docx");
    SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape> >()->LINQ_ToList();
    ASSERT_EQ(2, shapes->get_Count());

    // The two shapes are identical in terms of dimensions and shape type
    ASPOSE_ASSERT_EQ(shapes->idx_get(0)->get_Width(), shapes->idx_get(1)->get_Width());
    ASPOSE_ASSERT_EQ(shapes->idx_get(0)->get_Height(), shapes->idx_get(1)->get_Height());
    ASSERT_EQ(shapes->idx_get(0)->get_ShapeType(), shapes->idx_get(1)->get_ShapeType());

    // However, the first shape has no effects, while the second one has a shadow and thick outline
    ASPOSE_ASSERT_EQ(0.0, shapes->idx_get(0)->get_StrokeWeight());
    ASPOSE_ASSERT_EQ(20.0, shapes->idx_get(1)->get_StrokeWeight());
    ASSERT_FALSE(shapes->idx_get(0)->get_ShadowEnabled());
    ASSERT_TRUE(shapes->idx_get(1)->get_ShadowEnabled());

    // These effects make the size of the second shape's silhouette bigger than that of the first
    // Even though the size of the rectangle that shows up when we click on these shapes in Microsoft Word is the same,
    // the practical outer bounds of the second shape are affected by the shadow and outline and are bigger
    // We can use the AdjustWithEffects method to see exactly how much bigger they are

    // The first shape has no outline or effects
    SharedPtr<Shape> shape = shapes->idx_get(0);

    // Create a RectangleF object, which represents a rectangle, which we could potentially use as the coordinates and bounds for a shape
    System::Drawing::RectangleF rectangleF(200.0f, 200.0f, 1000.0f, 1000.0f);

    // Run this method to get the size of the rectangle adjusted for all of our shape's effects
    System::Drawing::RectangleF rectangleFOut = shape->AdjustWithEffects(rectangleF);

    // Since the shape has no border-changing effects, its boundary dimensions are unaffected
    ASPOSE_ASSERT_EQ(200, rectangleFOut.get_X());
    ASPOSE_ASSERT_EQ(200, rectangleFOut.get_Y());
    ASPOSE_ASSERT_EQ(1000, rectangleFOut.get_Width());
    ASPOSE_ASSERT_EQ(1000, rectangleFOut.get_Height());

    // The final extent of the first shape, in points
    ASPOSE_ASSERT_EQ(0, shape->get_BoundsWithEffects().get_X());
    ASPOSE_ASSERT_EQ(0, shape->get_BoundsWithEffects().get_Y());
    ASPOSE_ASSERT_EQ(147, shape->get_BoundsWithEffects().get_Width());
    ASPOSE_ASSERT_EQ(147, shape->get_BoundsWithEffects().get_Height());

    // Do the same with the second shape
    shape = shapes->idx_get(1);
    rectangleF = System::Drawing::RectangleF(200.0f, 200.0f, 1000.0f, 1000.0f);
    rectangleFOut = shape->AdjustWithEffects(rectangleF);

    // The shape's x/y coordinates (top left corner location) have been pushed back by the thick outline
    ASPOSE_ASSERT_EQ(171.5, rectangleFOut.get_X());
    ASPOSE_ASSERT_EQ(167, rectangleFOut.get_Y());

    // The width and height were also affected by the outline and shadow
    ASPOSE_ASSERT_EQ(1045, rectangleFOut.get_Width());
    ASPOSE_ASSERT_EQ(1132, rectangleFOut.get_Height());

    // These values are also affected by effects
    ASPOSE_ASSERT_EQ(-28.5, shape->get_BoundsWithEffects().get_X());
    ASPOSE_ASSERT_EQ(-33, shape->get_BoundsWithEffects().get_Y());
    ASPOSE_ASSERT_EQ(192, shape->get_BoundsWithEffects().get_Width());
    ASPOSE_ASSERT_EQ(279, shape->get_BoundsWithEffects().get_Height());
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
    //ExSummary:Shows how to export shapes to files in the local file system using a shape renderer.
    // Open a document that contains shapes and get its shape collection
    auto doc = MakeObject<Document>(MyDir + u"Various shapes.docx");
    SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape> >()->LINQ_ToList();
    ASSERT_EQ(7, shapes->get_Count());

    // There are 7 shapes in the document, with one group shape with 2 child shapes
    // The child shapes will be rendered but their parent group shape will be skipped, so we will see 6 output files
    for (auto shape : System::IterateOver(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape> >()))
    {
        SharedPtr<ShapeRenderer> renderer = shape->GetShapeRenderer();
        auto options = MakeObject<ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
        renderer->Save(ArtifactsDir + String::Format(u"Shape.RenderAllShapes.{0}.png",shape->get_Name()), options);
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
    //ExSummary:Shows how to detect that Shape has a SmartArt object.
    auto doc = MakeObject<Document>(MyDir + u"SmartArt.docx");

    int count = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape>>()->LINQ_Count([](SharedPtr<Shape> shape) { return shape->get_HasSmartArt(); });

    System::Console::WriteLine(u"The document has {0} shapes with SmartArt.", System::ObjectExt::Box<int>(count));
    //ExEnd

    ASSERT_EQ(2, count);
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
    //ExFor:OfficeMathRenderer.#ctor(Math.OfficeMath)
    //ExSummary:Shows how to measure and scale shapes.
    // Open a document that contains an OfficeMath object
    auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

    // Create a renderer for the OfficeMath object
    auto officeMath = System::DynamicCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    auto renderer = MakeObject<Aspose::Words::Rendering::OfficeMathRenderer>(officeMath);

    // We can measure the size of the image that the OfficeMath object will create when we render it
    ASSERT_NEAR(119.0f, renderer->get_SizeInPoints().get_Width(), 0.2f);
    ASSERT_NEAR(13.0f, renderer->get_SizeInPoints().get_Height(), 0.1f);

    ASSERT_NEAR(119.0f, renderer->get_BoundsInPoints().get_Width(), 0.2f);
    ASSERT_NEAR(13.0f, renderer->get_BoundsInPoints().get_Height(), 0.1f);

    // Shapes with transparent parts may return different values here
    ASSERT_NEAR(119.0f, renderer->get_OpaqueBoundsInPoints().get_Width(), 0.2f);
    ASSERT_NEAR(14.2f, renderer->get_OpaqueBoundsInPoints().get_Height(), 0.1f);

    // Get the shape size in pixels, with linear scaling to a specific DPI
    System::Drawing::Rectangle bounds = renderer->GetBoundsInPixels(1.0f, 96.0f);
    ASSERT_EQ(159, bounds.get_Width());
    ASSERT_EQ(18, bounds.get_Height());

    // Get the shape size in pixels, but with a different DPI for the horizontal and vertical dimensions
    bounds = renderer->GetBoundsInPixels(1.0f, 96.0f, 150.0f);
    ASSERT_EQ(159, bounds.get_Width());
    ASSERT_EQ(28, bounds.get_Height());

    // The opaque bounds may vary here also
    bounds = renderer->GetOpaqueBoundsInPixels(1.0f, 96.0f);
    ASSERT_EQ(159, bounds.get_Width());
    ASSERT_EQ(18, bounds.get_Height());

    bounds = renderer->GetOpaqueBoundsInPixels(1.0f, 96.0f, 150.0f);
    ASSERT_EQ(159, bounds.get_Width());
    ASSERT_EQ(30, bounds.get_Height());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExShape, OfficeMathRenderer)
{
    s_instance->OfficeMathRenderer();
}

} // namespace gtest_test

} // namespace ApiExamples
