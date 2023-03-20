#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/BorderCollection.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartTitle.h>
#include <Aspose.Words.Cpp/Drawing/DashStyle.h>
#include <Aspose.Words.Cpp/Drawing/EndCap.h>
#include <Aspose.Words.Cpp/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Drawing/FlipOrientation.h>
#include <Aspose.Words.Cpp/Drawing/GradientStop.h>
#include <Aspose.Words.Cpp/Drawing/GradientStopCollection.h>
#include <Aspose.Words.Cpp/Drawing/GradientStyle.h>
#include <Aspose.Words.Cpp/Drawing/GradientVariant.h>
#include <Aspose.Words.Cpp/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/JoinStyle.h>
#include <Aspose.Words.Cpp/Drawing/LayoutFlow.h>
#include <Aspose.Words.Cpp/Drawing/Ole/Forms2OleControl.h>
#include <Aspose.Words.Cpp/Drawing/Ole/Forms2OleControlCollection.h>
#include <Aspose.Words.Cpp/Drawing/Ole/Forms2OleControlType.h>
#include <Aspose.Words.Cpp/Drawing/Ole/OleControl.h>
#include <Aspose.Words.Cpp/Drawing/OleFormat.h>
#include <Aspose.Words.Cpp/Drawing/OlePackage.h>
#include <Aspose.Words.Cpp/Drawing/PatternType.h>
#include <Aspose.Words.Cpp/Drawing/PresetTexture.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeLineStyle.h>
#include <Aspose.Words.Cpp/Drawing/ShapeMarkupLanguage.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Drawing/SignatureLine.h>
#include <Aspose.Words.Cpp/Drawing/Stroke.h>
#include <Aspose.Words.Cpp/Drawing/TextBox.h>
#include <Aspose.Words.Cpp/Drawing/TextBoxAnchor.h>
#include <Aspose.Words.Cpp/Drawing/TextBoxWrapMode.h>
#include <Aspose.Words.Cpp/Drawing/TextPath.h>
#include <Aspose.Words.Cpp/Drawing/TextPathAlignment.h>
#include <Aspose.Words.Cpp/Drawing/TextureAlignment.h>
#include <Aspose.Words.Cpp/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Drawing/WrapSide.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/HeightRule.h>
#include <Aspose.Words.Cpp/LineStyle.h>
#include <Aspose.Words.Cpp/Math/MathObjectType.h>
#include <Aspose.Words.Cpp/Math/OfficeMath.h>
#include <Aspose.Words.Cpp/Math/OfficeMathDisplayType.h>
#include <Aspose.Words.Cpp/Math/OfficeMathJustification.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Rendering/OfficeMathRenderer.h>
#include <Aspose.Words.Cpp/Rendering/ShapeRenderer.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/SignatureLineOptions.h>
#include <Aspose.Words.Cpp/StoryType.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleType.h>
#include <Aspose.Words.Cpp/TableStyle.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Underline.h>
#include <Aspose.Words.Cpp/VisitorAction.h>
#include <drawing/color.h>
#include <drawing/image.h>
#include <drawing/point.h>
#include <drawing/point_f.h>
#include <drawing/rectangle.h>
#include <drawing/rectangle_f.h>
#include <drawing/size.h>
#include <drawing/size_f.h>
#include <system/collections/ienumerable.h>
#include <system/collections/list.h>
#include <system/details/dispose_guard.h>
#include <system/enum.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/guid.h>
#include <system/io/file.h>
#include <system/io/file_info.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/io/path.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/encoding.h>
#include <system/text/string_builder.h>
#include <system/type_info.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Ole;
using namespace Aspose::Words::Math;
using namespace Aspose::Words::Rendering;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
using namespace Aspose::Words::Tables;

namespace ApiExamples {

/// <summary>
/// Examples using shapes in documents.
/// </summary>
class ExShape : public ApiExampleBase
{
public:
    
    void Font(bool hideShape)
    {
        //ExStart
        //ExFor:ShapeBase.Font
        //ExFor:ShapeBase.ParentParagraph
        //ExSummary:Shows how to insert a text box, and set the font of its contents.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");

        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::TextBox, 300, 50);
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
            shape->get_Font()->set_Underline(Underline::Dash);
        }

        // Move the builder out of the text box back into the main document.
        builder->MoveTo(shape->get_ParentParagraph());

        builder->Writeln(u"\nThis text is outside the text box.");

        doc->Save(ArtifactsDir + u"Shape.Font.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.Font.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASPOSE_ASSERT_EQ(hideShape, shape->get_Font()->get_Hidden());

        if (hideShape)
        {
            ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), shape->get_Font()->get_HighlightColor().ToArgb());
            ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), shape->get_Font()->get_Color().ToArgb());
            ASSERT_EQ(Underline::None, shape->get_Font()->get_Underline());
        }
        else
        {
            ASSERT_EQ(System::Drawing::Color::get_Silver().ToArgb(), shape->get_Font()->get_HighlightColor().ToArgb());
            ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), shape->get_Font()->get_Color().ToArgb());
            ASSERT_EQ(Underline::Dash, shape->get_Font()->get_Underline());
        }

        TestUtil::VerifyShape(ShapeType::TextBox, u"TextBox 100002", 300.0, 50.0, 0, 0, shape);
        ASSERT_EQ(u"This text is inside the text box.", shape->GetText().Trim());
        ASSERT_EQ(u"Hello world!\rThis text is inside the text box.\r\rThis text is outside the text box.", doc->GetText().Trim());
    }

    void Rotate()
    {
        //ExStart
        //ExFor:ShapeBase.CanHaveImage
        //ExFor:ShapeBase.Rotation
        //ExSummary:Shows how to insert and rotate an image.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a shape with an image.
        SharedPtr<Shape> shape = builder->InsertImage(System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg"));
        ASSERT_TRUE(shape->get_CanHaveImage());
        ASSERT_TRUE(shape->get_HasImage());

        // Rotate the image 45 degrees clockwise.
        shape->set_Rotation(45);

        doc->Save(ArtifactsDir + u"Shape.Rotate.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.Rotate.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyShape(ShapeType::Image, String::Empty, 300.0, 300.0, 0, 0, shape);
        ASSERT_TRUE(shape->get_CanHaveImage());
        ASSERT_TRUE(shape->get_HasImage());
        ASPOSE_ASSERT_EQ(45.0, shape->get_Rotation());
    }

    void AspectRatioLockedDefaultValue()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromFile(ImageDir + u"Transparent background logo.png");

        SharedPtr<Shape> shape = builder->InsertImage(image);
        shape->set_WrapType(WrapType::None);
        shape->set_BehindText(true);

        shape->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);
        shape->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);

        // Calculate image left and top position so it appears in the center of the page.
        shape->set_Left((builder->get_PageSetup()->get_PageWidth() - shape->get_Width()) / 2);
        shape->set_Top((builder->get_PageSetup()->get_PageHeight() - shape->get_Height()) / 2);

        doc = DocumentHelper::SaveOpen(doc);

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        ASPOSE_ASSERT_EQ(true, shape->get_AspectRatioLocked());
    }

    void Coordinates()
    {
        //ExStart
        //ExFor:ShapeBase.DistanceBottom
        //ExFor:ShapeBase.DistanceLeft
        //ExFor:ShapeBase.DistanceRight
        //ExFor:ShapeBase.DistanceTop
        //ExSummary:Shows how to set the wrapping distance for a text that surrounds a shape.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a rectangle and, get the text to wrap tightly around its bounds.
        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Rectangle, 150, 150);
        shape->set_WrapType(WrapType::Tight);

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
        builder->Write(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ") +
                       u"Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        doc->Save(ArtifactsDir + u"Shape.Coordinates.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.Coordinates.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyShape(ShapeType::Rectangle, u"Rectangle 100002", 150.0, 150.0, 75.0, 150.0, shape);
        ASPOSE_ASSERT_EQ(40.0, shape->get_DistanceBottom());
        ASPOSE_ASSERT_EQ(40.0, shape->get_DistanceLeft());
        ASPOSE_ASSERT_EQ(40.0, shape->get_DistanceRight());
        ASPOSE_ASSERT_EQ(40.0, shape->get_DistanceTop());
        ASPOSE_ASSERT_EQ(60.0, shape->get_Rotation());
    }

    void GroupShape_()
    {
        //ExStart
        //ExFor:ShapeBase.Bounds
        //ExFor:ShapeBase.CoordOrigin
        //ExFor:ShapeBase.CoordSize
        //ExSummary:Shows how to create and populate a group shape.
        auto doc = MakeObject<Document>();

        // Create a group shape. A group shape can display a collection of child shape nodes.
        // In Microsoft Word, clicking within the group shape's boundary or on one of the group shape's child shapes will
        // select all the other child shapes within this group and allow us to scale and move all the shapes at once.
        auto group = MakeObject<GroupShape>(doc);

        ASSERT_EQ(WrapType::None, group->get_WrapType());

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
        auto rectangle = MakeObject<Shape>(doc, ShapeType::Rectangle);
        rectangle->set_Width(group->get_CoordSize().get_Width());
        rectangle->set_Height(group->get_CoordSize().get_Height());
        rectangle->set_Left(group->get_CoordOrigin().get_X());
        rectangle->set_Top(group->get_CoordOrigin().get_Y());
        group->AppendChild(rectangle);

        // Once a shape is a part of a group shape, we can access it as a child node and then modify it.
        (System::DynamicCast<Shape>(group->GetChild(NodeType::Shape, 0, true)))->get_Stroke()->set_DashStyle(DashStyle::Dash);

        // Create a small red star and insert it into the group.
        // Line up the shape with the group's coordinate origin, which we have moved to the center.
        auto smallRedStar = MakeObject<Shape>(doc, ShapeType::Star);
        smallRedStar->set_Width(20);
        smallRedStar->set_Height(20);
        smallRedStar->set_Left(-10);
        smallRedStar->set_Top(-10);
        smallRedStar->set_FillColor(System::Drawing::Color::get_Red());
        group->AppendChild(smallRedStar);

        // Insert a rectangle, and then insert a slightly smaller rectangle in the same place with an image.
        // Newer shapes that we add to the group overlap older shapes. The light blue rectangle will partially overlap the red star,
        // and then the shape with the image will overlap the light blue rectangle, using it as a frame.
        // We cannot use the "ZOrder" properties of shapes to manipulate their arrangement within a group shape.
        auto lightBlueRectangle = MakeObject<Shape>(doc, ShapeType::Rectangle);
        lightBlueRectangle->set_Width(250);
        lightBlueRectangle->set_Height(250);
        lightBlueRectangle->set_Left(-250);
        lightBlueRectangle->set_Top(-250);
        lightBlueRectangle->set_FillColor(System::Drawing::Color::get_LightBlue());
        group->AppendChild(lightBlueRectangle);

        auto image = MakeObject<Shape>(doc, ShapeType::Image);
        image->set_Width(200);
        image->set_Height(200);
        image->set_Left(-225);
        image->set_Top(-225);
        group->AppendChild(image);

        (System::DynamicCast<Shape>(group->GetChild(NodeType::Shape, 3, true)))->get_ImageData()->SetImage(ImageDir + u"Logo.jpg");

        // Insert a text box into the group shape. Set the "Left" property so that the text box's right edge
        // touches the right boundary of the group shape. Set the "Top" property so that the text box sits outside
        // the boundary of the group shape, with its top size lined up along the group shape's bottom margin.
        auto textBox = MakeObject<Shape>(doc, ShapeType::TextBox);
        textBox->set_Width(200);
        textBox->set_Height(50);
        textBox->set_Left(group->get_CoordSize().get_Width() + group->get_CoordOrigin().get_X() - 200);
        textBox->set_Top(group->get_CoordSize().get_Height() + group->get_CoordOrigin().get_Y());
        group->AppendChild(textBox);

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertNode(group);
        builder->MoveTo((System::DynamicCast<Shape>(group->GetChild(NodeType::Shape, 4, true)))->AppendChild(MakeObject<Paragraph>(doc)));
        builder->Write(u"Hello world!");

        doc->Save(ArtifactsDir + u"Shape.GroupShape.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.GroupShape.docx");
        group = System::DynamicCast<GroupShape>(doc->GetChild(NodeType::GroupShape, 0, true));

        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 0.0f, 400.0f, 400.0f), group->get_Bounds());
        ASPOSE_ASSERT_EQ(System::Drawing::Size(500, 500), group->get_CoordSize());
        ASPOSE_ASSERT_EQ(System::Drawing::Point(-250, -250), group->get_CoordOrigin());

        TestUtil::VerifyShape(ShapeType::Rectangle, String::Empty, 500.0, 500.0, -250.0, -250.0,
                              System::DynamicCast<Shape>(group->GetChild(NodeType::Shape, 0, true)));
        TestUtil::VerifyShape(ShapeType::Star, String::Empty, 20.0, 20.0, -10.0, -10.0, System::DynamicCast<Shape>(group->GetChild(NodeType::Shape, 1, true)));
        TestUtil::VerifyShape(ShapeType::Rectangle, String::Empty, 250.0, 250.0, -250.0, -250.0,
                              System::DynamicCast<Shape>(group->GetChild(NodeType::Shape, 2, true)));
        TestUtil::VerifyShape(ShapeType::Image, String::Empty, 200.0, 200.0, -225.0, -225.0,
                              System::DynamicCast<Shape>(group->GetChild(NodeType::Shape, 3, true)));
        TestUtil::VerifyShape(ShapeType::TextBox, String::Empty, 200.0, 50.0, 250.0, 50.0,
                              System::DynamicCast<Shape>(group->GetChild(NodeType::Shape, 4, true)));
    }

    void IsTopLevel()
    {
        //ExStart
        //ExFor:ShapeBase.IsTopLevel
        //ExSummary:Shows how to tell whether a shape is a part of a group shape.
        auto doc = MakeObject<Document>();

        auto shape = MakeObject<Shape>(doc, ShapeType::Rectangle);
        shape->set_Width(200);
        shape->set_Height(200);
        shape->set_WrapType(WrapType::None);

        // A shape by default is not part of any group shape, and therefore has the "IsTopLevel" property set to "true".
        ASSERT_TRUE(shape->get_IsTopLevel());

        auto group = MakeObject<GroupShape>(doc);
        group->AppendChild(shape);

        // Once we assimilate a shape into a group shape, the "IsTopLevel" property changes to "false".
        ASSERT_FALSE(shape->get_IsTopLevel());
        //ExEnd
    }

    void LocalToParent()
    {
        //ExStart
        //ExFor:ShapeBase.CoordOrigin
        //ExFor:ShapeBase.CoordSize
        //ExFor:ShapeBase.LocalToParent(PointF)
        //ExSummary:Shows how to translate the x and y coordinate location on a shape's coordinate plane to a location on the parent shape's coordinate plane.
        auto doc = MakeObject<Document>();

        // Insert a group shape, and place it 100 points below and to the right of
        // the document's x and Y coordinate origin point.
        auto group = MakeObject<GroupShape>(doc);
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

        auto shape = MakeObject<Shape>(doc, ShapeType::Rectangle);
        shape->set_Width(100);
        shape->set_Height(100);
        shape->set_Left(700);
        shape->set_Top(700);

        group->AppendChild(shape);
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(group);

        doc->Save(ArtifactsDir + u"Shape.LocalToParent.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.LocalToParent.docx");
        group = System::DynamicCast<GroupShape>(doc->GetChild(NodeType::GroupShape, 0, true));

        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(100.0f, 100.0f, 500.0f, 500.0f), group->get_Bounds());
        ASPOSE_ASSERT_EQ(System::Drawing::Size(500, 500), group->get_CoordSize());
        ASPOSE_ASSERT_EQ(System::Drawing::Point(-250, -250), group->get_CoordOrigin());
    }

    void AnchorLocked(bool anchorLocked)
    {
        //ExStart
        //ExFor:ShapeBase.AnchorLocked
        //ExSummary:Shows how to lock or unlock a shape's paragraph anchor.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");

        builder->Write(u"Our shape will have an anchor attached to this paragraph.");
        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Rectangle, 200, 160);
        shape->set_WrapType(WrapType::None);
        builder->InsertBreak(BreakType::ParagraphBreak);

        builder->Writeln(u"Hello again!");

        // Set the "AnchorLocked" property to "true" to prevent the shape's anchor
        // from moving when moving the shape in Microsoft Word.
        // Set the "AnchorLocked" property to "false" to allow any movement of the shape
        // to also move its anchor to any other paragraph that the shape ends up close to.
        shape->set_AnchorLocked(anchorLocked);

        // If the shape does not have a visible anchor symbol to its left,
        // we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
        doc->Save(ArtifactsDir + u"Shape.AnchorLocked.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.AnchorLocked.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASPOSE_ASSERT_EQ(anchorLocked, shape->get_AnchorLocked());
    }

    void DeleteAllShapes()
    {
        //ExStart
        //ExFor:Shape
        //ExSummary:Shows how to delete all shapes from a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert two shapes along with a group shape with another shape inside it.
        builder->InsertShape(ShapeType::Rectangle, 400, 200);
        builder->InsertShape(ShapeType::Star, 300, 300);

        auto group = MakeObject<GroupShape>(doc);
        group->set_Bounds(System::Drawing::RectangleF(100.0f, 50.0f, 200.0f, 100.0f));
        group->set_CoordOrigin(System::Drawing::Point(-1000, -500));

        auto subShape = MakeObject<Shape>(doc, ShapeType::Cube);
        subShape->set_Width(500);
        subShape->set_Height(700);
        subShape->set_Left(0);
        subShape->set_Top(0);

        group->AppendChild(subShape);
        builder->InsertNode(group);

        ASSERT_EQ(3, doc->GetChildNodes(NodeType::Shape, true)->get_Count());
        ASSERT_EQ(1, doc->GetChildNodes(NodeType::GroupShape, true)->get_Count());

        // Remove all Shape nodes from the document.
        SharedPtr<NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);
        shapes->Clear();

        // All shapes are gone, but the group shape is still in the document.
        ASSERT_EQ(1, doc->GetChildNodes(NodeType::GroupShape, true)->get_Count());
        ASSERT_EQ(0, doc->GetChildNodes(NodeType::Shape, true)->get_Count());

        // Remove all group shapes separately.
        SharedPtr<NodeCollection> groupShapes = doc->GetChildNodes(NodeType::GroupShape, true);
        groupShapes->Clear();

        ASSERT_EQ(0, doc->GetChildNodes(NodeType::GroupShape, true)->get_Count());
        ASSERT_EQ(0, doc->GetChildNodes(NodeType::Shape, true)->get_Count());
        //ExEnd
    }

    void IsInline()
    {
        //ExStart
        //ExFor:ShapeBase.IsInline
        //ExSummary:Shows how to determine whether a shape is inline or floating.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are two wrapping types that shapes may have.
        // 1 -  Inline:
        builder->Write(u"Hello world! ");
        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Rectangle, 100, 100);
        shape->set_FillColor(System::Drawing::Color::get_LightBlue());
        builder->Write(u" Hello again.");

        // An inline shape sits inside a paragraph among other paragraph elements, such as runs of text.
        // In Microsoft Word, we may click and drag the shape to any paragraph as if it is a character.
        // If the shape is large, it will affect vertical paragraph spacing.
        // We cannot move this shape to a place with no paragraph.
        ASSERT_EQ(WrapType::Inline, shape->get_WrapType());
        ASSERT_TRUE(shape->get_IsInline());

        // 2 -  Floating:
        shape = builder->InsertShape(ShapeType::Rectangle, RelativeHorizontalPosition::LeftMargin, 200, RelativeVerticalPosition::TopMargin, 200, 100, 100,
                                     WrapType::None);
        shape->set_FillColor(System::Drawing::Color::get_Orange());

        // A floating shape belongs to the paragraph that we insert it into,
        // which we can determine by an anchor symbol that appears when we click the shape.
        // If the shape does not have a visible anchor symbol to its left,
        // we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
        // In Microsoft Word, we may left click and drag this shape freely to any location.
        ASSERT_EQ(WrapType::None, shape->get_WrapType());
        ASSERT_FALSE(shape->get_IsInline());

        doc->Save(ArtifactsDir + u"Shape.IsInline.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.IsInline.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyShape(ShapeType::Rectangle, u"Rectangle 100002", 100, 100, 0, 0, shape);
        ASSERT_EQ(System::Drawing::Color::get_LightBlue().ToArgb(), shape->get_FillColor().ToArgb());
        ASSERT_EQ(WrapType::Inline, shape->get_WrapType());
        ASSERT_TRUE(shape->get_IsInline());

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        TestUtil::VerifyShape(ShapeType::Rectangle, u"Rectangle 100004", 100, 100, 200, 200, shape);
        ASSERT_EQ(System::Drawing::Color::get_Orange().ToArgb(), shape->get_FillColor().ToArgb());
        ASSERT_EQ(WrapType::None, shape->get_WrapType());
        ASSERT_FALSE(shape->get_IsInline());
    }

    void Bounds()
    {
        //ExStart
        //ExFor:ShapeBase.Bounds
        //ExFor:ShapeBase.BoundsInPoints
        //ExSummary:Shows how to verify shape containing block boundaries.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> line = builder->InsertShape(ShapeType::Line, RelativeHorizontalPosition::LeftMargin, 50, RelativeVerticalPosition::TopMargin, 50, 100,
                                                     100, WrapType::None);
        line->set_StrokeColor(System::Drawing::Color::get_Orange());

        // Even though the line itself takes up little space on the document page,
        // it occupies a rectangular containing block, the size of which we can determine using the "Bounds" properties.
        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(50.0f, 50.0f, 100.0f, 100.0f), line->get_Bounds());
        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(50.0f, 50.0f, 100.0f, 100.0f), line->get_BoundsInPoints());

        // Create a group shape, and then set the size of its containing block using the "Bounds" property.
        auto group = MakeObject<GroupShape>(doc);
        group->set_Bounds(System::Drawing::RectangleF(0.0f, 100.0f, 250.0f, 250.0f));

        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 100.0f, 250.0f, 250.0f), group->get_BoundsInPoints());

        // Create a rectangle, verify the size of its bounding block, and then add it to the group shape.
        auto rectangle = MakeObject<Shape>(doc, ShapeType::Rectangle);
        rectangle->set_Width(100);
        rectangle->set_Height(100);
        rectangle->set_Left(700);
        rectangle->set_Top(700);

        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(700.0f, 700.0f, 100.0f, 100.0f), rectangle->get_BoundsInPoints());

        group->AppendChild(rectangle);

        // The group shape's coordinate plane has its origin on the top left-hand side corner of its containing block,
        // and the x and y coordinates of (1000, 1000) on the bottom right-hand side corner.
        // Our group shape is 250x250pt in size, so every 4pt on the group shape's coordinate plane
        // translates to 1pt in the document body's coordinate plane.
        // Every shape that we insert will also shrink in size by a factor of 4.
        // The change in the shape's "BoundsInPoints" property will reflect this.
        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(175.0f, 275.0f, 25.0f, 25.0f), rectangle->get_BoundsInPoints());

        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(group);

        // Insert a shape and place it outside of the bounds of the group shape's containing block.
        auto shape = MakeObject<Shape>(doc, ShapeType::Rectangle);
        shape->set_Width(100);
        shape->set_Height(100);
        shape->set_Left(1000);
        shape->set_Top(1000);

        group->AppendChild(shape);

        // The group shape's footprint in the document body has increased, but the containing block remains the same.
        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 100.0f, 250.0f, 250.0f), group->get_BoundsInPoints());
        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(250.0f, 350.0f, 25.0f, 25.0f), shape->get_BoundsInPoints());

        doc->Save(ArtifactsDir + u"Shape.Bounds.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.Bounds.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyShape(ShapeType::Line, u"Line 100002", 100, 100, 50, 50, shape);
        ASSERT_EQ(System::Drawing::Color::get_Orange().ToArgb(), shape->get_StrokeColor().ToArgb());
        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(50.0f, 50.0f, 100.0f, 100.0f), shape->get_BoundsInPoints());

        group = System::DynamicCast<GroupShape>(doc->GetChild(NodeType::GroupShape, 0, true));

        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 100.0f, 250.0f, 250.0f), group->get_Bounds());
        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(0.0f, 100.0f, 250.0f, 250.0f), group->get_BoundsInPoints());
        ASPOSE_ASSERT_EQ(System::Drawing::Size(1000, 1000), group->get_CoordSize());
        ASPOSE_ASSERT_EQ(System::Drawing::Point(0, 0), group->get_CoordOrigin());

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        TestUtil::VerifyShape(ShapeType::Rectangle, String::Empty, 100, 100, 700, 700, shape);
        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(175.0f, 275.0f, 25.0f, 25.0f), shape->get_BoundsInPoints());

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 2, true));

        TestUtil::VerifyShape(ShapeType::Rectangle, String::Empty, 100, 100, 1000, 1000, shape);
        ASPOSE_ASSERT_EQ(System::Drawing::RectangleF(250.0f, 350.0f, 25.0f, 25.0f), shape->get_BoundsInPoints());
    }

    void FlipShapeOrientation()
    {
        //ExStart
        //ExFor:ShapeBase.FlipOrientation
        //ExFor:FlipOrientation
        //ExSummary:Shows how to flip a shape on an axis.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert an image shape and leave its orientation in its default state.
        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Rectangle, RelativeHorizontalPosition::LeftMargin, 100, RelativeVerticalPosition::TopMargin,
                                                      100, 100, 100, WrapType::None);
        shape->get_ImageData()->SetImage(ImageDir + u"Logo.jpg");

        ASSERT_EQ(FlipOrientation::None, shape->get_FlipOrientation());

        shape = builder->InsertShape(ShapeType::Rectangle, RelativeHorizontalPosition::LeftMargin, 250, RelativeVerticalPosition::TopMargin, 100, 100, 100,
                                     WrapType::None);
        shape->get_ImageData()->SetImage(ImageDir + u"Logo.jpg");

        // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the second shape on the y-axis,
        // making it into a horizontal mirror image of the first shape.
        shape->set_FlipOrientation(FlipOrientation::Horizontal);

        shape = builder->InsertShape(ShapeType::Rectangle, RelativeHorizontalPosition::LeftMargin, 100, RelativeVerticalPosition::TopMargin, 250, 100, 100,
                                     WrapType::None);
        shape->get_ImageData()->SetImage(ImageDir + u"Logo.jpg");

        // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the third shape on the x-axis,
        // making it into a vertical mirror image of the first shape.
        shape->set_FlipOrientation(FlipOrientation::Vertical);

        shape = builder->InsertShape(ShapeType::Rectangle, RelativeHorizontalPosition::LeftMargin, 250, RelativeVerticalPosition::TopMargin, 250, 100, 100,
                                     WrapType::None);
        shape->get_ImageData()->SetImage(ImageDir + u"Logo.jpg");

        // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the fourth shape on both the x and y axes,
        // making it into a horizontal and vertical mirror image of the first shape.
        shape->set_FlipOrientation(FlipOrientation::Both);

        doc->Save(ArtifactsDir + u"Shape.FlipShapeOrientation.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.FlipShapeOrientation.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyShape(ShapeType::Rectangle, u"Rectangle 100002", 100, 100, 100, 100, shape);
        ASSERT_EQ(FlipOrientation::None, shape->get_FlipOrientation());

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        TestUtil::VerifyShape(ShapeType::Rectangle, u"Rectangle 100004", 100, 100, 100, 250, shape);
        ASSERT_EQ(FlipOrientation::Horizontal, shape->get_FlipOrientation());

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 2, true));

        TestUtil::VerifyShape(ShapeType::Rectangle, u"Rectangle 100006", 100, 100, 250, 100, shape);
        ASSERT_EQ(FlipOrientation::Vertical, shape->get_FlipOrientation());

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 3, true));

        TestUtil::VerifyShape(ShapeType::Rectangle, u"Rectangle 100008", 100, 100, 250, 250, shape);
        ASSERT_EQ(FlipOrientation::Both, shape->get_FlipOrientation());
    }

    void TextureFill()
    {
        //ExStart
        //ExFor:Fill.TextureAlignment
        //ExFor:TextureAlignment
        //ExSummary:Shows how to fill and tiling the texture inside the shape.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Rectangle, 80, 80);

        // Apply texture alignment to the shape fill.
        shape->get_Fill()->PresetTextured(PresetTexture::Canvas);
        shape->get_Fill()->set_TextureAlignment(TextureAlignment::TopRight);

        // Use the compliance option to define the shape using DML if you want to get "TextureAlignment"
        // property after the document saves.
        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_Compliance(OoxmlCompliance::Iso29500_2008_Strict);

        doc->Save(ArtifactsDir + u"Shape.TextureFill.docx", saveOptions);
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.TextureFill.docx");

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_EQ(TextureAlignment::TopRight, shape->get_Fill()->get_TextureAlignment());
    }

    void GradientFill()
    {
        //ExStart
        //ExFor:Fill.OneColorGradient(Color, GradientStyle, GradientVariant, Double)
        //ExFor:Fill.OneColorGradient(GradientStyle, GradientVariant, Double)
        //ExFor:Fill.TwoColorGradient(Color, Color, GradientStyle, GradientVariant)
        //ExFor:Fill.TwoColorGradient(GradientStyle, GradientVariant)
        //ExFor:Fill.GradientStyle
        //ExFor:Fill.GradientVariant
        //ExFor:Fill.GradientAngle
        //ExFor:GradientStyle
        //ExFor:GradientVariant
        //ExSummary:Shows how to fill a shape with a gradients.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Rectangle, 80, 80);
        // Apply One-color gradient fill to the shape with ForeColor of gradient fill.
        shape->get_Fill()->OneColorGradient(System::Drawing::Color::get_Red(), GradientStyle::Horizontal, GradientVariant::Variant2, 0.1);

        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), shape->get_Fill()->get_ForeColor().ToArgb());
        ASSERT_EQ(GradientStyle::Horizontal, shape->get_Fill()->get_GradientStyle());
        ASSERT_EQ(GradientVariant::Variant2, shape->get_Fill()->get_GradientVariant());
        ASPOSE_ASSERT_EQ(270, shape->get_Fill()->get_GradientAngle());

        shape = builder->InsertShape(ShapeType::Rectangle, 80, 80);
        // Apply Two-color gradient fill to the shape.
        shape->get_Fill()->TwoColorGradient(GradientStyle::FromCorner, GradientVariant::Variant4);
        // Change BackColor of gradient fill.
        shape->get_Fill()->set_BackColor(System::Drawing::Color::get_Yellow());
        // Note that changes "GradientAngle" for "GradientStyle.FromCorner/GradientStyle.FromCenter"
        // gradient fill don't get any effect, it will work only for linear gradient.
        shape->get_Fill()->set_GradientAngle(15);

        ASSERT_EQ(System::Drawing::Color::get_Yellow().ToArgb(), shape->get_Fill()->get_BackColor().ToArgb());
        ASSERT_EQ(GradientStyle::FromCorner, shape->get_Fill()->get_GradientStyle());
        ASSERT_EQ(GradientVariant::Variant4, shape->get_Fill()->get_GradientVariant());
        ASPOSE_ASSERT_EQ(0, shape->get_Fill()->get_GradientAngle());

        // Use the compliance option to define the shape using DML if you want to get "GradientStyle",
        // "GradientVariant" and "GradientAngle" properties after the document saves.
        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_Compliance(OoxmlCompliance::Iso29500_2008_Strict);

        doc->Save(ArtifactsDir + u"Shape.GradientFill.docx", saveOptions);
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.GradientFill.docx");
        auto firstShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), firstShape->get_Fill()->get_ForeColor().ToArgb());
        ASSERT_EQ(GradientStyle::Horizontal, firstShape->get_Fill()->get_GradientStyle());
        ASSERT_EQ(GradientVariant::Variant2, firstShape->get_Fill()->get_GradientVariant());
        ASPOSE_ASSERT_EQ(270, firstShape->get_Fill()->get_GradientAngle());

        auto secondShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        ASSERT_EQ(System::Drawing::Color::get_Yellow().ToArgb(), secondShape->get_Fill()->get_BackColor().ToArgb());
        ASSERT_EQ(GradientStyle::FromCorner, secondShape->get_Fill()->get_GradientStyle());
        ASSERT_EQ(GradientVariant::Variant4, secondShape->get_Fill()->get_GradientVariant());
        ASPOSE_ASSERT_EQ(0, secondShape->get_Fill()->get_GradientAngle());
    }

    void GradientStops()
    {
        //ExStart
        //ExFor:Fill.GradientStops
        //ExFor:GradientStopCollection
        //ExFor:GradientStopCollection.Insert(System.Int32, GradientStop)
        //ExFor:GradientStopCollection.Add(GradientStop)
        //ExFor:GradientStopCollection.RemoveAt(System.Int32)
        //ExFor:GradientStopCollection.Remove(GradientStop)
        //ExFor:GradientStopCollection.Item(System.Int32)
        //ExFor:GradientStopCollection.Count
        //ExFor:GradientStop.#ctor(Color, Double)
        //ExFor:GradientStop.#ctor(Color, Double, Double)
        //ExFor:GradientStop.Color
        //ExFor:GradientStop.Position
        //ExFor:GradientStop.Transparency
        //ExFor:GradientStop.Remove
        //ExSummary:Shows how to add gradient stops to the gradient fill.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Rectangle, 80, 80);
        shape->get_Fill()->TwoColorGradient(System::Drawing::Color::get_Green(), System::Drawing::Color::get_Red(), GradientStyle::Horizontal,
                                            GradientVariant::Variant2);

        // Get gradient stops collection.
        SharedPtr<GradientStopCollection> gradientStops = shape->get_Fill()->get_GradientStops();

        // Change first gradient stop.
        gradientStops->idx_get(0)->set_Color(System::Drawing::Color::get_Aqua());
        gradientStops->idx_get(0)->set_Position(0.1);
        gradientStops->idx_get(0)->set_Transparency(0.25);

        // Add new gradient stop to the end of collection.
        auto gradientStop = MakeObject<GradientStop>(System::Drawing::Color::get_Brown(), 0.5);
        gradientStops->Add(gradientStop);

        // Remove gradient stop at index 1.
        gradientStops->RemoveAt(1);
        // And insert new gradient stop at the same index 1.
        gradientStops->Insert(1, MakeObject<GradientStop>(System::Drawing::Color::get_Chocolate(), 0.75, 0.3));

        // Remove last gradient stop in the collection.
        gradientStop = gradientStops->idx_get(2);
        gradientStops->Remove(gradientStop);

        ASSERT_EQ(2, gradientStops->get_Count());

        ASSERT_EQ(System::Drawing::Color::get_Aqua().ToArgb(), gradientStops->idx_get(0)->get_Color().ToArgb());
        ASSERT_NEAR(0.1, gradientStops->idx_get(0)->get_Position(), 0.01);
        ASSERT_NEAR(0.25, gradientStops->idx_get(0)->get_Transparency(), 0.01);

        ASSERT_EQ(System::Drawing::Color::get_Chocolate().ToArgb(), gradientStops->idx_get(1)->get_Color().ToArgb());
        ASSERT_NEAR(0.75, gradientStops->idx_get(1)->get_Position(), 0.01);
        ASSERT_NEAR(0.3, gradientStops->idx_get(1)->get_Transparency(), 0.01);

        // Use the compliance option to define the shape using DML
        // if you want to get "GradientStops" property after the document saves.
        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_Compliance(OoxmlCompliance::Iso29500_2008_Strict);

        doc->Save(ArtifactsDir + u"Shape.GradientStops.docx", saveOptions);
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.GradientStops.docx");

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        gradientStops = shape->get_Fill()->get_GradientStops();

        ASSERT_EQ(2, gradientStops->get_Count());

        ASSERT_EQ(System::Drawing::Color::get_Aqua().ToArgb(), gradientStops->idx_get(0)->get_Color().ToArgb());
        ASSERT_NEAR(0.1, gradientStops->idx_get(0)->get_Position(), 0.01);
        ASSERT_NEAR(0.25, gradientStops->idx_get(0)->get_Transparency(), 0.01);

        ASSERT_EQ(System::Drawing::Color::get_Chocolate().ToArgb(), gradientStops->idx_get(1)->get_Color().ToArgb());
        ASSERT_NEAR(0.75, gradientStops->idx_get(1)->get_Position(), 0.01);
        ASSERT_NEAR(0.3, gradientStops->idx_get(1)->get_Transparency(), 0.01);
    }

    void FillPattern()
    {
        //ExStart
        //ExFor:Fill.Patterned(PatternType)
        //ExFor:Fill.Patterned(PatternType, Color, Color)
        //ExSummary:Shows how to set pattern for a shape.
        auto doc = MakeObject<Document>(MyDir + u"Shape stroke pattern border.docx");

        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        SharedPtr<Fill> fill = shape->get_Fill();

        std::cout << String::Format(u"Pattern value is: {0}", fill->get_Pattern()) << std::endl;

        // There are several ways specified fill to a pattern.
        // 1 -  Apply pattern to the shape fill:
        fill->Patterned(PatternType::DiagonalBrick);

        // 2 -  Apply pattern with foreground and background colors to the shape fill:
        fill->Patterned(PatternType::DiagonalBrick, System::Drawing::Color::get_Aqua(), System::Drawing::Color::get_Bisque());

        doc->Save(ArtifactsDir + u"Shape.FillPattern.docx");
        //ExEnd
    }

    void Title()
    {
        //ExStart
        //ExFor:ShapeBase.Title
        //ExSummary:Shows how to set the title of a shape.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a shape, give it a title, and then add it to the document.
        auto shape = MakeObject<Shape>(doc, ShapeType::Cube);
        shape->set_Width(200);
        shape->set_Height(200);
        shape->set_Title(u"My cube");

        builder->InsertNode(shape);

        // When we save a document with a shape that has a title,
        // Aspose.Words will store that title in the shape's Alt Text.
        doc->Save(ArtifactsDir + u"Shape.Title.docx");

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.Title.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_EQ(String::Empty, shape->get_Title());
        ASSERT_EQ(u"Title: My cube", shape->get_AlternativeText());
        //ExEnd

        TestUtil::VerifyShape(ShapeType::Cube, String::Empty, 200.0, 200.0, 0.0, 0.0, shape);
    }

    void ReplaceTextboxesWithImages()
    {
        //ExStart
        //ExFor:WrapSide
        //ExFor:ShapeBase.WrapSide
        //ExFor:NodeCollection
        //ExFor:CompositeNode.InsertAfter(Node, Node)
        //ExFor:NodeCollection.ToArray
        //ExSummary:Shows how to replace all textbox shapes with image shapes.
        auto doc = MakeObject<Document>(MyDir + u"Textboxes in drawing canvas.docx");

        ArrayPtr<SharedPtr<Shape>> shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_ToArray();

        auto isTextBox = [](SharedPtr<Shape> s)
        {
            return s->get_ShapeType() == ShapeType::TextBox;
        };
        auto isImage = [](SharedPtr<Shape> s)
        {
            return s->get_ShapeType() == ShapeType::Image;
        };
        ASSERT_EQ(3, shapes->LINQ_Count(isTextBox));
        ASSERT_EQ(1, shapes->LINQ_Count(isImage));

        for (SharedPtr<Shape> shape : shapes)
        {
            if (shape->get_ShapeType() == ShapeType::TextBox)
            {
                auto replacementShape = MakeObject<Shape>(doc, ShapeType::Image);
                replacementShape->get_ImageData()->SetImage(ImageDir + u"Logo.jpg");
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

                shape->get_ParentNode()->InsertAfter(replacementShape, shape);
                shape->Remove();
            }
        }

        shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_ToArray();

        ASSERT_EQ(0, shapes->LINQ_Count(isTextBox));
        ASSERT_EQ(4, shapes->LINQ_Count(isImage));

        doc->Save(ArtifactsDir + u"Shape.ReplaceTextboxesWithImages.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.ReplaceTextboxesWithImages.docx");
        auto outShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_EQ(WrapSide::Both, outShape->get_WrapSide());
    }

    void CreateTextBox()
    {
        //ExStart
        //ExFor:Shape.#ctor(DocumentBase, ShapeType)
        //ExFor:Story.FirstParagraph
        //ExFor:Shape.FirstParagraph
        //ExFor:ShapeBase.WrapType
        //ExSummary:Shows how to create and format a text box.
        auto doc = MakeObject<Document>();

        // Create a floating text box.
        auto textBox = MakeObject<Shape>(doc, ShapeType::TextBox);
        textBox->set_WrapType(WrapType::None);
        textBox->set_Height(50);
        textBox->set_Width(200);

        // Set the horizontal, and vertical alignment of the text inside the shape.
        textBox->set_HorizontalAlignment(HorizontalAlignment::Center);
        textBox->set_VerticalAlignment(VerticalAlignment::Top);

        // Add a paragraph to the text box and add a run of text that the text box will display.
        textBox->AppendChild(MakeObject<Paragraph>(doc));
        SharedPtr<Paragraph> para = textBox->get_FirstParagraph();
        para->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
        auto run = MakeObject<Run>(doc);
        run->set_Text(u"Hello world!");
        para->AppendChild(run);

        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(textBox);

        doc->Save(ArtifactsDir + u"Shape.CreateTextBox.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.CreateTextBox.docx");
        textBox = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyShape(ShapeType::TextBox, String::Empty, 200.0, 50.0, 0.0, 0.0, textBox);
        ASSERT_EQ(WrapType::None, textBox->get_WrapType());
        ASSERT_EQ(HorizontalAlignment::Center, textBox->get_HorizontalAlignment());
        ASSERT_EQ(VerticalAlignment::Top, textBox->get_VerticalAlignment());
        ASSERT_EQ(u"Hello world!", textBox->GetText().Trim());
    }

    void ZOrder()
    {
        //ExStart
        //ExFor:ShapeBase.ZOrder
        //ExSummary:Shows how to manipulate the order of shapes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert three different colored rectangles that partially overlap each other.
        // When we insert a shape that overlaps another shape, Aspose.Words places the newer shape on top of the old one.
        // The light green rectangle will overlap the light blue rectangle and partially obscure it,
        // and the light blue rectangle will obscure the orange rectangle.
        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Rectangle, RelativeHorizontalPosition::LeftMargin, 100, RelativeVerticalPosition::TopMargin,
                                                      100, 200, 200, WrapType::None);
        shape->set_FillColor(System::Drawing::Color::get_Orange());

        shape = builder->InsertShape(ShapeType::Rectangle, RelativeHorizontalPosition::LeftMargin, 150, RelativeVerticalPosition::TopMargin, 150, 200, 200,
                                     WrapType::None);
        shape->set_FillColor(System::Drawing::Color::get_LightBlue());

        shape = builder->InsertShape(ShapeType::Rectangle, RelativeHorizontalPosition::LeftMargin, 200, RelativeVerticalPosition::TopMargin, 200, 200, 200,
                                     WrapType::None);
        shape->set_FillColor(System::Drawing::Color::get_LightGreen());

        ArrayPtr<SharedPtr<Shape>> shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_ToArray();

        // The "ZOrder" property of a shape determines its stacking priority among other overlapping shapes.
        // If two overlapping shapes have different "ZOrder" values,
        // Microsoft Word will place the shape with a higher value over the shape with the lower value.
        // Set the "ZOrder" values of our shapes to place the first orange rectangle over the second light blue one
        // and the second light blue rectangle over the third light green rectangle.
        // This will reverse their original stacking order.
        shapes[0]->set_ZOrder(3);
        shapes[1]->set_ZOrder(2);
        shapes[2]->set_ZOrder(1);

        doc->Save(ArtifactsDir + u"Shape.ZOrder.docx");
        //ExEnd
    }

    void GetOleObjectRawData()
    {
        //ExStart
        //ExFor:OleFormat.GetRawData
        //ExSummary:Shows how to access the raw data of an embedded OLE object.
        auto doc = MakeObject<Document>(MyDir + u"OLE objects.docx");

        for (const auto& shape : System::IterateOver(doc->GetChildNodes(NodeType::Shape, true)))
        {
            SharedPtr<OleFormat> oleFormat = (System::DynamicCast<Shape>(shape))->get_OleFormat();
            if (oleFormat != nullptr)
            {
                std::cout << "This is " << (oleFormat->get_IsLink() ? String(u"a linked") : String(u"an embedded")) << " object" << std::endl;
                ArrayPtr<uint8_t> oleRawData = oleFormat->GetRawData();

                ASSERT_EQ(24576, oleRawData->get_Length());
            }
        }
        //ExEnd
    }

    void OleControl_()
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
        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        // The OLE object in the first shape is a Microsoft Excel spreadsheet.
        SharedPtr<OleFormat> oleFormat = shape->get_OleFormat();

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
            auto fs = MakeObject<System::IO::FileStream>(ArtifactsDir + u"OLE spreadsheet extracted via stream" + oleFormat->get_SuggestedExtension(),
                                                         System::IO::FileMode::Create);
            oleFormat->Save(fs);
        }

        // 2 -  Save it directly to a filename:
        oleFormat->Save(ArtifactsDir + u"OLE spreadsheet saved directly" + oleFormat->get_SuggestedExtension());
        //ExEnd

        ASSERT_LT(8000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"OLE spreadsheet extracted via stream.xlsx")->get_Length());
        ASSERT_LT(8000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"OLE spreadsheet saved directly.xlsx")->get_Length());
    }

    void OleLinks()
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

        // Embed a Microsoft Visio drawing into the document as an OLE object.
        builder->InsertOleObject(ImageDir + u"Microsoft Visio drawing.vsd", u"Package", false, false, nullptr);

        // Insert a link to the file in the local file system and display it as an icon.
        builder->InsertOleObject(ImageDir + u"Microsoft Visio drawing.vsd", u"Package", true, true, nullptr);

        // Inserting OLE objects creates shapes that store these objects.
        ArrayPtr<SharedPtr<Shape>> shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_ToArray();

        ASSERT_EQ(2, shapes->get_Length());
        auto isOleObject = [](SharedPtr<Shape> s)
        {
            return s->get_ShapeType() == ShapeType::OleObject;
        };
        ASSERT_EQ(2, shapes->LINQ_Count(isOleObject));

        // If a shape contains an OLE object, it will have a valid "OleFormat" property,
        // which we can use to verify some aspects of the shape.
        SharedPtr<OleFormat> oleFormat = shapes[0]->get_OleFormat();

        ASPOSE_ASSERT_EQ(false, oleFormat->get_IsLink());
        ASPOSE_ASSERT_EQ(false, oleFormat->get_OleIcon());

        oleFormat = shapes[1]->get_OleFormat();

        ASPOSE_ASSERT_EQ(true, oleFormat->get_IsLink());
        ASPOSE_ASSERT_EQ(true, oleFormat->get_OleIcon());

        ASSERT_TRUE(oleFormat->get_SourceFullName().EndsWith(String(u"Images") + System::IO::Path::DirectorySeparatorChar + u"Microsoft Visio drawing.vsd"));
        ASSERT_EQ(u"", oleFormat->get_SourceItem());

        ASSERT_EQ(u"Microsoft Visio drawing.vsd", oleFormat->get_IconCaption());

        doc->Save(ArtifactsDir + u"Shape.OleLinks.docx");

        // If the object contains OLE data, we can access it using a stream.
        {
            SharedPtr<System::IO::MemoryStream> stream = oleFormat->GetOleEntry(u"\x0001"
                                                                                u"CompObj");
            ArrayPtr<uint8_t> oleEntryBytes = stream->ToArray();
            ASSERT_EQ(76, oleEntryBytes->get_Length());
        }
        //ExEnd
    }

    void OleControlCollection()
    {
        //ExStart
        //ExFor:OleFormat.Clsid
        //ExFor:Ole.Forms2OleControlCollection
        //ExFor:Ole.Forms2OleControlCollection.Count
        //ExFor:Ole.Forms2OleControlCollection.Item(Int32)
        //ExSummary:Shows how to access an OLE control embedded in a document and its child controls.
        auto doc = MakeObject<Document>(MyDir + u"OLE ActiveX controls.docm");

        // Shapes store and display OLE objects in the document's body.
        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_EQ(u"6e182020-f460-11ce-9bcd-00aa00608e01", System::ObjectExt::ToString(shape->get_OleFormat()->get_Clsid()));

        auto oleControl = System::DynamicCast<Forms2OleControl>(shape->get_OleFormat()->get_OleControl());

        // Some OLE controls may contain child controls, such as the one in this document with three options buttons.
        SharedPtr<Forms2OleControlCollection> oleControlCollection = oleControl->get_ChildNodes();

        ASSERT_EQ(3, oleControlCollection->get_Count());

        ASSERT_EQ(u"C#", oleControlCollection->idx_get(0)->get_Caption());
        ASSERT_EQ(u"1", oleControlCollection->idx_get(0)->get_Value());

        ASSERT_EQ(u"Visual Basic", oleControlCollection->idx_get(1)->get_Caption());
        ASSERT_EQ(u"0", oleControlCollection->idx_get(1)->get_Value());

        ASSERT_EQ(u"Delphi", oleControlCollection->idx_get(2)->get_Caption());
        ASSERT_EQ(u"0", oleControlCollection->idx_get(2)->get_Value());
        //ExEnd
    }

    void SuggestedFileName()
    {
        //ExStart
        //ExFor:OleFormat.SuggestedFileName
        //ExSummary:Shows how to get an OLE object's suggested file name.
        auto doc = MakeObject<Document>(MyDir + u"OLE shape.rtf");

        auto oleShape = System::DynamicCast<Shape>(doc->get_FirstSection()->get_Body()->GetChild(NodeType::Shape, 0, true));

        // OLE objects can provide a suggested filename and extension,
        // which we can use when saving the object's contents into a file in the local file system.
        String suggestedFileName = oleShape->get_OleFormat()->get_SuggestedFileName();

        ASSERT_EQ(u"CSV.csv", suggestedFileName);

        {
            auto fileStream = MakeObject<System::IO::FileStream>(ArtifactsDir + suggestedFileName, System::IO::FileMode::Create);
            oleShape->get_OleFormat()->Save(fileStream);
        }
        //ExEnd
    }

    void ObjectDidNotHaveSuggestedFileName()
    {
        auto doc = MakeObject<Document>(MyDir + u"ActiveX controls.docx");

        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        ASSERT_EQ(shape->get_OleFormat()->get_SuggestedFileName(), String::Empty);
    }

    void ResolutionDefaultValues()
    {
        auto imageOptions = MakeObject<ImageSaveOptions>(SaveFormat::Jpeg);

        ASPOSE_ASSERT_EQ(96, imageOptions->get_HorizontalResolution());
        ASPOSE_ASSERT_EQ(96, imageOptions->get_VerticalResolution());
    }

    void RenderOfficeMath()
    {
        //ExStart
        //ExFor:ImageSaveOptions.Scale
        //ExFor:OfficeMath.GetMathRenderer
        //ExFor:NodeRendererBase.Save(String, ImageSaveOptions)
        //ExSummary:Shows how to render an Office Math object into an image file in the local file system.
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

        auto math = System::DynamicCast<OfficeMath>(doc->GetChild(NodeType::OfficeMath, 0, true));

        // Create an "ImageSaveOptions" object to pass to the node renderer's "Save" method to modify
        // how it renders the OfficeMath node into an image.
        auto saveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Png);

        // Set the "Scale" property to 5 to render the object to five times its original size.
        saveOptions->set_Scale(5.0f);

        math->GetMathRenderer()->Save(ArtifactsDir + u"Shape.RenderOfficeMath.png", saveOptions);
        //ExEnd

        if (!IsRunningOnMono())
        {
            TestUtil::VerifyImage(795, 87, ArtifactsDir + u"Shape.RenderOfficeMath.png");
        }
        else
        {
            TestUtil::VerifyImage(735, 128, ArtifactsDir + u"Shape.RenderOfficeMath.png");
        }
    }

    void OfficeMathDisplayException()
    {
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

        auto officeMath = System::DynamicCast<OfficeMath>(doc->GetChild(NodeType::OfficeMath, 0, true));
        officeMath->set_DisplayType(OfficeMathDisplayType::Display);

        ASSERT_THROW(officeMath->set_Justification(OfficeMathJustification::Inline), System::ArgumentException);
    }

    void OfficeMathDefaultValue()
    {
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

        auto officeMath = System::DynamicCast<OfficeMath>(doc->GetChild(NodeType::OfficeMath, 6, true));

        ASSERT_EQ(OfficeMathDisplayType::Inline, officeMath->get_DisplayType());
        ASSERT_EQ(OfficeMathJustification::Inline, officeMath->get_Justification());
    }

    void OfficeMath_()
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

        auto officeMath = System::DynamicCast<OfficeMath>(doc->GetChild(NodeType::OfficeMath, 0, true));

        // OfficeMath nodes that are children of other OfficeMath nodes are always inline.
        // The node we are working with is the base node to change its location and display type.
        ASSERT_EQ(MathObjectType::OMathPara, officeMath->get_MathObjectType());
        ASSERT_EQ(NodeType::OfficeMath, officeMath->get_NodeType());
        ASPOSE_ASSERT_EQ(officeMath->get_ParentNode(), officeMath->get_ParentParagraph());

        // OOXML and WML formats use the "EquationXmlEncoding" property.
        ASSERT_TRUE(officeMath->get_EquationXmlEncoding() == nullptr);

        // Change the location and display type of the OfficeMath node.
        officeMath->set_DisplayType(OfficeMathDisplayType::Display);
        officeMath->set_Justification(OfficeMathJustification::Left);

        doc->Save(ArtifactsDir + u"Shape.OfficeMath.docx");
        //ExEnd

        ASSERT_TRUE(DocumentHelper::CompareDocs(ArtifactsDir + u"Shape.OfficeMath.docx", GoldsDir + u"Shape.OfficeMath Gold.docx"));
    }

    void CannotBeSetDisplayWithInlineJustification()
    {
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

        auto officeMath = System::DynamicCast<OfficeMath>(doc->GetChild(NodeType::OfficeMath, 0, true));
        officeMath->set_DisplayType(OfficeMathDisplayType::Display);

        ASSERT_THROW(officeMath->set_Justification(OfficeMathJustification::Inline), System::ArgumentException);
    }

    void CannotBeSetInlineDisplayWithJustification()
    {
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

        auto officeMath = System::DynamicCast<OfficeMath>(doc->GetChild(NodeType::OfficeMath, 0, true));
        officeMath->set_DisplayType(OfficeMathDisplayType::Inline);

        ASSERT_THROW(officeMath->set_Justification(OfficeMathJustification::Center), System::ArgumentException);
    }

    void OfficeMathDisplayNestedObjects()
    {
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

        auto officeMath = System::DynamicCast<OfficeMath>(doc->GetChild(NodeType::OfficeMath, 0, true));

        ASSERT_EQ(OfficeMathDisplayType::Display, officeMath->get_DisplayType());
        ASSERT_EQ(OfficeMathJustification::Center, officeMath->get_Justification());
    }

    void WorkWithMathObjectType(int index, MathObjectType objectType)
    {
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

        auto officeMath = System::DynamicCast<OfficeMath>(doc->GetChild(NodeType::OfficeMath, index, true));
        ASSERT_EQ(objectType, officeMath->get_MathObjectType());
    }

    void AspectRatio(bool lockAspectRatio)
    {
        //ExStart
        //ExFor:ShapeBase.AspectRatioLocked
        //ExSummary:Shows how to lock/unlock a shape's aspect ratio.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a shape. If we open this document in Microsoft Word, we can left click the shape to reveal
        // eight sizing handles around its perimeter, which we can click and drag to change its size.
        SharedPtr<Shape> shape = builder->InsertImage(ImageDir + u"Logo.jpg");

        // Set the "AspectRatioLocked" property to "true" to preserve the shape's aspect ratio
        // when using any of the four diagonal sizing handles, which change both the image's height and width.
        // Using any orthogonal sizing handles that either change the height or width will still change the aspect ratio.
        // Set the "AspectRatioLocked" property to "false" to allow us to
        // freely change the image's aspect ratio with all sizing handles.
        shape->set_AspectRatioLocked(lockAspectRatio);

        doc->Save(ArtifactsDir + u"Shape.AspectRatio.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.AspectRatio.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASPOSE_ASSERT_EQ(lockAspectRatio, shape->get_AspectRatioLocked());
    }

    void MarkupLanguageByDefault()
    {
        //ExStart
        //ExFor:ShapeBase.MarkupLanguage
        //ExFor:ShapeBase.SizeInPoints
        //ExSummary:Shows how to verify a shape's size and markup language.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertImage(ImageDir + u"Transparent background logo.png");

        ASSERT_EQ(ShapeMarkupLanguage::Dml, shape->get_MarkupLanguage());
        ASPOSE_ASSERT_EQ(System::Drawing::SizeF(300.0f, 300.0f), shape->get_SizeInPoints());
        //ExEnd
    }

    void MarkupLunguageForDifferentMsWordVersions(MsWordVersion msWordVersion, ShapeMarkupLanguage shapeMarkupLanguage)
    {
        auto doc = MakeObject<Document>();
        doc->get_CompatibilityOptions()->OptimizeFor(msWordVersion);

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertImage(ImageDir + u"Transparent background logo.png");

        for (const auto& shape : System::IterateOver(doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()))
        {
            ASSERT_EQ(shapeMarkupLanguage, shape->get_MarkupLanguage());
        }
    }

    void Stroke_()
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

        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Rectangle, RelativeHorizontalPosition::LeftMargin, 100, RelativeVerticalPosition::TopMargin,
                                                      100, 200, 200, WrapType::None);

        // Basic shapes, such as the rectangle, have two visible parts.
        // 1 -  The fill, which applies to the area within the outline of the shape:
        shape->get_Fill()->set_ForeColor(System::Drawing::Color::get_White());

        // 2 -  The stroke, which marks the outline of the shape:
        // Modify various properties of this shape's stroke.
        SharedPtr<Stroke> stroke = shape->get_Stroke();
        stroke->set_On(true);
        stroke->set_Weight(5);
        stroke->set_Color(System::Drawing::Color::get_Red());
        stroke->set_DashStyle(DashStyle::ShortDashDotDot);
        stroke->set_JoinStyle(JoinStyle::Miter);
        stroke->set_EndCap(EndCap::Square);
        stroke->set_LineStyle(ShapeLineStyle::Triple);

        doc->Save(ArtifactsDir + u"Shape.Stroke.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.Stroke.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        stroke = shape->get_Stroke();

        ASPOSE_ASSERT_EQ(true, stroke->get_On());
        ASPOSE_ASSERT_EQ(5, stroke->get_Weight());
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), stroke->get_Color().ToArgb());
        ASSERT_EQ(DashStyle::ShortDashDotDot, stroke->get_DashStyle());
        ASSERT_EQ(JoinStyle::Miter, stroke->get_JoinStyle());
        ASSERT_EQ(EndCap::Square, stroke->get_EndCap());
        ASSERT_EQ(ShapeLineStyle::Triple, stroke->get_LineStyle());
    }

    void InsertOleObjectAsHtmlFile()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertOleObject(u"http://www.aspose.com", u"htmlfile", true, false, nullptr);

        doc->Save(ArtifactsDir + u"Shape.InsertOleObjectAsHtmlFile.docx");
    }

    void InsertOlePackage()
    {
        //ExStart
        //ExFor:OlePackage
        //ExFor:OleFormat.OlePackage
        //ExFor:OlePackage.FileName
        //ExFor:OlePackage.DisplayName
        //ExSummary:Shows how insert an OLE object into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // OLE objects allow us to open other files in the local file system using another installed application
        // in our operating system by double-clicking on the shape that contains the OLE object in the document body.
        // In this case, our external file will be a ZIP archive.
        ArrayPtr<uint8_t> zipFileBytes = System::IO::File::ReadAllBytes(DatabaseDir + u"cat001.zip");

        {
            auto stream = MakeObject<System::IO::MemoryStream>(zipFileBytes);
            SharedPtr<Shape> shape = builder->InsertOleObject(stream, u"Package", true, nullptr);

            shape->get_OleFormat()->get_OlePackage()->set_FileName(u"Package file name.zip");
            shape->get_OleFormat()->get_OlePackage()->set_DisplayName(u"Package display name.zip");
        }

        doc->Save(ArtifactsDir + u"Shape.InsertOlePackage.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.InsertOlePackage.docx");
        auto getShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_EQ(u"Package file name.zip", getShape->get_OleFormat()->get_OlePackage()->get_FileName());
        ASSERT_EQ(u"Package display name.zip", getShape->get_OleFormat()->get_OlePackage()->get_DisplayName());
    }

    void GetAccessToOlePackage()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> oleObject = builder->InsertOleObject(MyDir + u"Spreadsheet.xlsx", false, false, nullptr);
        SharedPtr<Shape> oleObjectAsOlePackage = builder->InsertOleObject(MyDir + u"Spreadsheet.xlsx", u"Excel.Sheet", false, false, nullptr);

        ASPOSE_ASSERT_EQ(nullptr, oleObject->get_OleFormat()->get_OlePackage());
        ASPOSE_ASSERT_EQ(System::ObjectExt::GetType<OlePackage>(), System::ObjectExt::GetType(oleObjectAsOlePackage->get_OleFormat()->get_OlePackage()));
    }

    void Resize()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Rectangle, 200, 300);
        shape->set_Height(300);
        shape->set_Width(500);
        shape->set_Rotation(30);

        doc->Save(ArtifactsDir + u"Shape.Resize.docx");
    }

    void Calendar()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartTable();
        builder->get_RowFormat()->set_Height(100);
        builder->get_RowFormat()->set_HeightRule(HeightRule::Exactly);

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

        SharedPtr<NodeCollection> runs = doc->GetChildNodes(NodeType::Run, true);
        int num = 1;

        for (const auto& run : System::IterateOver(runs->LINQ_OfType<SharedPtr<Run>>()))
        {
            auto watermark = MakeObject<Shape>(doc, ShapeType::TextPlainText);
            watermark->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);
            watermark->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);
            watermark->set_Width(30);
            watermark->set_Height(30);
            watermark->set_HorizontalAlignment(HorizontalAlignment::Center);
            watermark->set_VerticalAlignment(VerticalAlignment::Center);
            watermark->set_Rotation(-40);

            watermark->get_Fill()->set_ForeColor(System::Drawing::Color::get_Gainsboro());
            watermark->set_StrokeColor(System::Drawing::Color::get_Gainsboro());

            watermark->get_TextPath()->set_Text(String::Format(u"{0}", num));
            watermark->get_TextPath()->set_FontFamily(u"Arial");

            watermark->set_Name(String::Format(u"Watermark_{0}", num++));

            watermark->set_BehindText(true);

            builder->MoveTo(run);
            builder->InsertNode(watermark);
        }

        doc->Save(ArtifactsDir + u"Shape.Calendar.docx");

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.Calendar.docx");
        auto shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape>>()->LINQ_ToList();

        ASSERT_EQ(31, shapes->get_Count());

        for (const auto& shape : shapes)
        {
            TestUtil::VerifyShape(ShapeType::TextPlainText, String::Format(u"Watermark_{0}", shapes->IndexOf(shape) + 1), 30.0, 30.0, 0.0, 0.0, shape);
        }
    }

    void IsLayoutInCell(bool isLayoutInCell)
    {
        //ExStart
        //ExFor:ShapeBase.IsLayoutInCell
        //ExSummary:Shows how to determine how to display a shape in a table cell.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->InsertCell();
        builder->EndTable();

        auto tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->Add(StyleType::Table, u"MyTableStyle1"));
        tableStyle->set_BottomPadding(20);
        tableStyle->set_LeftPadding(10);
        tableStyle->set_RightPadding(10);
        tableStyle->set_TopPadding(20);
        tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Black());
        tableStyle->get_Borders()->set_LineStyle(LineStyle::Single);

        table->set_Style(tableStyle);

        builder->MoveTo(table->get_FirstRow()->get_FirstCell()->get_FirstParagraph());

        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Rectangle, RelativeHorizontalPosition::LeftMargin, 50, RelativeVerticalPosition::TopMargin,
                                                      100, 100, 100, WrapType::None);

        // Set the "IsLayoutInCell" property to "true" to display the shape as an inline element inside the cell's paragraph.
        // The coordinate origin that will determine the shape's location will be the top left corner of the shape's cell.
        // If we re-size the cell, the shape will move to maintain the same position starting from the cell's top left.
        // Set the "IsLayoutInCell" property to "false" to display the shape as an independent floating shape.
        // The coordinate origin that will determine the shape's location will be the top left corner of the page,
        // and the shape will not respond to any re-sizing of its cell.
        shape->set_IsLayoutInCell(isLayoutInCell);

        // We can only apply the "IsLayoutInCell" property to floating shapes.
        shape->set_WrapType(WrapType::None);

        doc->Save(ArtifactsDir + u"Shape.LayoutInTableCell.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.LayoutInTableCell.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        shape = System::DynamicCast<Shape>(table->get_FirstRow()->get_FirstCell()->GetChild(NodeType::Shape, 0, true));

        ASPOSE_ASSERT_EQ(isLayoutInCell, shape->get_IsLayoutInCell());
    }

    void ShapeInsertion()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertShape(ShapeType, RelativeHorizontalPosition, double, RelativeVerticalPosition, double, double, double, WrapType)
        //ExFor:DocumentBuilder.InsertShape(ShapeType, double, double)
        //ExFor:OoxmlCompliance
        //ExFor:OoxmlSaveOptions.Compliance
        //ExSummary:Shows how to insert DML shapes into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are two wrapping types that shapes may have.
        // 1 -  Floating:
        builder->InsertShape(ShapeType::TopCornersRounded, RelativeHorizontalPosition::Page, 100, RelativeVerticalPosition::Page, 100, 50, 50, WrapType::None);

        // 2 -  Inline:
        builder->InsertShape(ShapeType::DiagonalCornersRounded, 50, 50);

        // If you need to create "non-primitive" shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
        // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, or DiagonalCornersRounded,
        // then save the document with "Strict" or "Transitional" compliance, which allows saving shape as DML.
        auto saveOptions = MakeObject<OoxmlSaveOptions>(SaveFormat::Docx);
        saveOptions->set_Compliance(OoxmlCompliance::Iso29500_2008_Transitional);

        doc->Save(ArtifactsDir + u"Shape.ShapeInsertion.docx", saveOptions);
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.ShapeInsertion.docx");
        auto shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape>>()->LINQ_ToList();

        TestUtil::VerifyShape(ShapeType::TopCornersRounded, u"TopCornersRounded 100002", 50.0, 50.0, 100.0, 100.0, shapes->idx_get(0));
        TestUtil::VerifyShape(ShapeType::DiagonalCornersRounded, u"DiagonalCornersRounded 100004", 50.0, 50.0, 0.0, 0.0, shapes->idx_get(1));
    }

    //ExStart
    //ExFor:Shape.Accept(DocumentVisitor)
    //ExFor:Shape.Chart
    //ExFor:Shape.ExtrusionEnabled
    //ExFor:Shape.Filled
    //ExFor:Shape.HasChart
    //ExFor:Shape.OleFormat
    //ExFor:Shape.ShadowEnabled
    //ExFor:Shape.StoryType
    //ExFor:Shape.StrokeColor
    //ExFor:Shape.Stroked
    //ExFor:Shape.StrokeWeight
    //ExSummary:Shows how to iterate over all the shapes in a document.
    void VisitShapes()
    {
        auto doc = MakeObject<Document>(MyDir + u"Revision shape.docx");
        ASSERT_EQ(2, doc->GetChildNodes(NodeType::Shape, true)->get_Count());
        //ExSkip

        auto visitor = MakeObject<ExShape::ShapeAppearancePrinter>();
        doc->Accept(visitor);

        std::cout << visitor->GetText() << std::endl;
    }

    /// <summary>
    /// Logs appearance-related information about visited shapes.
    /// </summary>
    class ShapeAppearancePrinter : public DocumentVisitor
    {
    public:
        ShapeAppearancePrinter() : mShapesVisited(0), mTextIndentLevel(0)
        {
            mShapesVisited = 0;
            mTextIndentLevel = 0;
            mStringBuilder = MakeObject<System::Text::StringBuilder>();
        }

        /// <summary>
        /// Return all the text that the StringBuilder has accumulated.
        /// </summary>
        String GetText()
        {
            return String::Format(u"Shapes visited: {0}\n{1}", mShapesVisited, mStringBuilder);
        }

        /// <summary>
        /// Called when this visitor visits the start of a Shape node.
        /// </summary>
        VisitorAction VisitShapeStart(SharedPtr<Shape> shape) override
        {
            AppendLine(String::Format(u"Shape found: {0}", shape->get_ShapeType()));

            mTextIndentLevel++;

            if (shape->get_HasChart())
            {
                AppendLine(String::Format(u"Has chart: {0}", shape->get_Chart()->get_Title()->get_Text()));
            }

            AppendLine(String::Format(u"Extrusion enabled: {0}", shape->get_ExtrusionEnabled()));
            AppendLine(String::Format(u"Shadow enabled: {0}", shape->get_ShadowEnabled()));
            AppendLine(String::Format(u"StoryType: {0}", shape->get_StoryType()));

            if (shape->get_Stroked())
            {
                ASPOSE_EXPECT_EQ(shape->get_Stroke()->get_Color(), shape->get_StrokeColor());
                AppendLine(String::Format(u"Stroke colors: {0}, {1}", shape->get_Stroke()->get_Color(), shape->get_Stroke()->get_Color2()));
                AppendLine(String::Format(u"Stroke weight: {0}", shape->get_StrokeWeight()));
            }

            if (shape->get_Filled())
            {
                AppendLine(String::Format(u"Filled: {0}", shape->get_FillColor()));
            }

            if (shape->get_OleFormat() != nullptr)
            {
                AppendLine(String::Format(u"Ole found of type: {0}", shape->get_OleFormat()->get_ProgId()));
            }

            if (shape->get_SignatureLine() != nullptr)
            {
                AppendLine(String::Format(u"Found signature line for: {0}, {1}", shape->get_SignatureLine()->get_Signer(),
                                          shape->get_SignatureLine()->get_SignerTitle()));
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when this visitor visits the end of a Shape node.
        /// </summary>
        VisitorAction VisitShapeEnd(SharedPtr<Shape> shape) override
        {
            mTextIndentLevel--;
            mShapesVisited++;
            AppendLine(String::Format(u"End of {0}", shape->get_ShapeType()));

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when this visitor visits the start of a GroupShape node.
        /// </summary>
        VisitorAction VisitGroupShapeStart(SharedPtr<GroupShape> groupShape) override
        {
            AppendLine(String::Format(u"Shape group found: {0}", groupShape->get_ShapeType()));
            mTextIndentLevel++;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when this visitor visits the end of a GroupShape node.
        /// </summary>
        VisitorAction VisitGroupShapeEnd(SharedPtr<GroupShape> groupShape) override
        {
            mTextIndentLevel--;
            AppendLine(String::Format(u"End of {0}", groupShape->get_ShapeType()));

            return VisitorAction::Continue;
        }

    private:
        int mShapesVisited;
        int mTextIndentLevel;
        SharedPtr<System::Text::StringBuilder> mStringBuilder;

        /// <summary>
        /// Appends a line to the StringBuilder with one prepended tab character for each indent level.
        /// </summary>
        void AppendLine(String text)
        {
            for (int i = 0; i < mTextIndentLevel; i++)
            {
                mStringBuilder->Append(u'\t');
            }

            mStringBuilder->AppendLine(text);
        }
    };
    //ExEnd

    void SignatureLine_()
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
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto options = MakeObject<SignatureLineOptions>();
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
        SharedPtr<Shape> shape = builder->InsertSignatureLine(options, RelativeHorizontalPosition::RightMargin, -170.0, RelativeVerticalPosition::BottomMargin,
                                                              -60.0, WrapType::None);

        ASSERT_TRUE(shape->get_IsSignatureLine());

        // Verify the properties of our signature line via its Shape object.
        SharedPtr<SignatureLine> signatureLine = shape->get_SignatureLine();

        ASSERT_EQ(u"john.doe@management.com", signatureLine->get_Email());
        ASSERT_EQ(u"John Doe", signatureLine->get_Signer());
        ASSERT_EQ(u"Senior Manager", signatureLine->get_SignerTitle());
        ASSERT_EQ(u"Please sign here", signatureLine->get_Instructions());
        ASSERT_TRUE(signatureLine->get_ShowDate());
        ASSERT_TRUE(signatureLine->get_AllowComments());
        ASSERT_TRUE(signatureLine->get_DefaultInstructions());

        doc->Save(ArtifactsDir + u"Shape.SignatureLine.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.SignatureLine.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyShape(ShapeType::Image, String::Empty, 192.75, 96.75, -60.0, -170.0, shape);
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

    void TextBoxFitShapeToText()
    {
        //ExStart
        //ExFor:TextBox
        //ExFor:TextBox.FitShapeToText
        //ExSummary:Shows how to get a text box to resize itself to fit its contents tightly.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> textBoxShape = builder->InsertShape(ShapeType::TextBox, 150, 100);
        SharedPtr<TextBox> textBox = textBoxShape->get_TextBox();

        // Apply these values to both these members to get the parent shape to fit
        // tightly around the text contents, ignoring the dimensions we have set.
        textBox->set_FitShapeToText(true);
        textBox->set_TextBoxWrapMode(TextBoxWrapMode::None);

        builder->MoveTo(textBoxShape->get_LastParagraph());
        builder->Write(u"Text fit tightly inside textbox.");

        doc->Save(ArtifactsDir + u"Shape.TextBoxFitShapeToText.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.TextBoxFitShapeToText.docx");
        textBoxShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyShape(ShapeType::TextBox, u"TextBox 100002", 150.0, 100.0, 0.0, 0.0, textBoxShape);
        TestUtil::VerifyTextBox(LayoutFlow::Horizontal, true, TextBoxWrapMode::None, 3.6, 3.6, 7.2, 7.2, textBoxShape->get_TextBox());
        ASSERT_EQ(u"Text fit tightly inside textbox.", textBoxShape->GetText().Trim());
    }

    void TextBoxMargins()
    {
        //ExStart
        //ExFor:TextBox
        //ExFor:TextBox.InternalMarginBottom
        //ExFor:TextBox.InternalMarginLeft
        //ExFor:TextBox.InternalMarginRight
        //ExFor:TextBox.InternalMarginTop
        //ExSummary:Shows how to set internal margins for a text box.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert another textbox with specific margins.
        SharedPtr<Shape> textBoxShape = builder->InsertShape(ShapeType::TextBox, 100, 100);
        SharedPtr<TextBox> textBox = textBoxShape->get_TextBox();
        textBox->set_InternalMarginTop(15);
        textBox->set_InternalMarginBottom(15);
        textBox->set_InternalMarginLeft(15);
        textBox->set_InternalMarginRight(15);

        builder->MoveTo(textBoxShape->get_LastParagraph());
        builder->Write(u"Text placed according to textbox margins.");

        doc->Save(ArtifactsDir + u"Shape.TextBoxMargins.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.TextBoxMargins.docx");
        textBoxShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyShape(ShapeType::TextBox, u"TextBox 100002", 100.0, 100.0, 0.0, 0.0, textBoxShape);
        TestUtil::VerifyTextBox(LayoutFlow::Horizontal, false, TextBoxWrapMode::Square, 15.0, 15.0, 15.0, 15.0, textBoxShape->get_TextBox());
        ASSERT_EQ(u"Text placed according to textbox margins.", textBoxShape->GetText().Trim());
    }

    void TextBoxContentsWrapMode(TextBoxWrapMode textBoxWrapMode)
    {
        //ExStart
        //ExFor:TextBox.TextBoxWrapMode
        //ExFor:TextBoxWrapMode
        //ExSummary:Shows how to set a wrapping mode for the contents of a text box.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> textBoxShape = builder->InsertShape(ShapeType::TextBox, 300, 300);
        SharedPtr<TextBox> textBox = textBoxShape->get_TextBox();

        // Set the "TextBoxWrapMode" property to "TextBoxWrapMode.None" to increase the text box's width
        // to accommodate text, should it be large enough.
        // Set the "TextBoxWrapMode" property to "TextBoxWrapMode.Square" to
        // wrap all text inside the text box, preserving its dimensions.
        textBox->set_TextBoxWrapMode(textBoxWrapMode);

        builder->MoveTo(textBoxShape->get_LastParagraph());
        builder->get_Font()->set_Size(32);
        builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        doc->Save(ArtifactsDir + u"Shape.TextBoxContentsWrapMode.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.TextBoxContentsWrapMode.docx");
        textBoxShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyShape(ShapeType::TextBox, u"TextBox 100002", 300.0, 300.0, 0.0, 0.0, textBoxShape);
        TestUtil::VerifyTextBox(LayoutFlow::Horizontal, false, textBoxWrapMode, 3.6, 3.6, 7.2, 7.2, textBoxShape->get_TextBox());
        ASSERT_EQ(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
                  textBoxShape->GetText().Trim());
    }

    void TextBoxShapeType()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set compatibility options to correctly using of VerticalAnchor property.
        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2016);

        SharedPtr<Shape> textBoxShape = builder->InsertShape(ShapeType::TextBox, 100, 100);
        // Not all formats are compatible with this one.
        // For most of the incompatible formats, AW generated warnings on save, so use doc.WarningCallback to check it.
        textBoxShape->get_TextBox()->set_VerticalAnchor(TextBoxAnchor::Bottom);

        builder->MoveTo(textBoxShape->get_LastParagraph());
        builder->Write(u"Text placed bottom");

        doc->Save(ArtifactsDir + u"Shape.TextBoxShapeType.docx");
    }

    void CreateLinkBetweenTextBoxes()
    {
        //ExStart
        //ExFor:TextBox.IsValidLinkTarget(TextBox)
        //ExFor:TextBox.Next
        //ExFor:TextBox.Previous
        //ExFor:TextBox.BreakForwardLink
        //ExSummary:Shows how to link text boxes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> textBoxShape1 = builder->InsertShape(ShapeType::TextBox, 100, 100);
        SharedPtr<TextBox> textBox1 = textBoxShape1->get_TextBox();
        builder->Writeln();

        SharedPtr<Shape> textBoxShape2 = builder->InsertShape(ShapeType::TextBox, 100, 100);
        SharedPtr<TextBox> textBox2 = textBoxShape2->get_TextBox();
        builder->Writeln();

        SharedPtr<Shape> textBoxShape3 = builder->InsertShape(ShapeType::TextBox, 100, 100);
        SharedPtr<TextBox> textBox3 = textBoxShape3->get_TextBox();
        builder->Writeln();

        SharedPtr<Shape> textBoxShape4 = builder->InsertShape(ShapeType::TextBox, 100, 100);
        SharedPtr<TextBox> textBox4 = textBoxShape4->get_TextBox();

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

        doc->Save(ArtifactsDir + u"Shape.CreateLinkBetweenTextBoxes.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.CreateLinkBetweenTextBoxes.docx");
        auto shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_ToList();

        TestUtil::VerifyShape(ShapeType::TextBox, u"TextBox 100002", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(0));
        TestUtil::VerifyTextBox(LayoutFlow::Horizontal, false, TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(0)->get_TextBox());
        ASSERT_EQ(String::Empty, shapes->idx_get(0)->GetText().Trim());

        TestUtil::VerifyShape(ShapeType::TextBox, u"TextBox 100004", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(1));
        TestUtil::VerifyTextBox(LayoutFlow::Horizontal, false, TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(1)->get_TextBox());
        ASSERT_EQ(String::Empty, shapes->idx_get(1)->GetText().Trim());

        TestUtil::VerifyShape(ShapeType::Rectangle, u"TextBox 100006", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(2));
        TestUtil::VerifyTextBox(LayoutFlow::Horizontal, false, TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(2)->get_TextBox());
        ASSERT_EQ(String::Empty, shapes->idx_get(2)->GetText().Trim());

        TestUtil::VerifyShape(ShapeType::TextBox, u"TextBox 100008", 100.0, 100.0, 0.0, 0.0, shapes->idx_get(3));
        TestUtil::VerifyTextBox(LayoutFlow::Horizontal, false, TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shapes->idx_get(3)->get_TextBox());
        ASSERT_EQ(u"Hello world!", shapes->idx_get(3)->GetText().Trim());
    }

    void VerticalAnchor(TextBoxAnchor verticalAnchor)
    {
        //ExStart
        //ExFor:CompatibilityOptions
        //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
        //ExFor:TextBoxAnchor
        //ExFor:TextBox.VerticalAnchor
        //ExSummary:Shows how to vertically align the text contents of a text box.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::TextBox, 200, 200);

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
        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2007);
        doc->Save(ArtifactsDir + u"Shape.VerticalAnchor.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Shape.VerticalAnchor.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyShape(ShapeType::TextBox, u"TextBox 100002", 200.0, 200.0, 0.0, 0.0, shape);
        TestUtil::VerifyTextBox(LayoutFlow::Horizontal, false, TextBoxWrapMode::Square, 3.6, 3.6, 7.2, 7.2, shape->get_TextBox());
        ASSERT_EQ(verticalAnchor, shape->get_TextBox()->get_VerticalAnchor());
        ASSERT_EQ(u"Hello world!", shape->GetText().Trim());
    }

    //ExStart
    //ExFor:Shape.TextPath
    //ExFor:ShapeBase.IsWordArt
    //ExFor:TextPath
    //ExFor:TextPath.Bold
    //ExFor:TextPath.FitPath
    //ExFor:TextPath.FitShape
    //ExFor:TextPath.FontFamily
    //ExFor:TextPath.Italic
    //ExFor:TextPath.Kerning
    //ExFor:TextPath.On
    //ExFor:TextPath.ReverseRows
    //ExFor:TextPath.RotateLetters
    //ExFor:TextPath.SameLetterHeights
    //ExFor:TextPath.Shadow
    //ExFor:TextPath.SmallCaps
    //ExFor:TextPath.Spacing
    //ExFor:TextPath.StrikeThrough
    //ExFor:TextPath.Text
    //ExFor:TextPath.TextPathAlignment
    //ExFor:TextPath.Trim
    //ExFor:TextPath.Underline
    //ExFor:TextPath.XScale
    //ExFor:TextPathAlignment
    //ExSummary:Shows how to work with WordArt.
    void InsertTextPaths()
    {
        auto doc = MakeObject<Document>();

        // Insert a WordArt object to display text in a shape that we can re-size and move by using the mouse in Microsoft Word.
        // Provide a "ShapeType" as an argument to set a shape for the WordArt.
        SharedPtr<Shape> shape = AppendWordArt(doc, u"Hello World! This text is bold, and italic.", u"Arial", 480, 24, System::Drawing::Color::get_White(),
                                               System::Drawing::Color::get_Black(), ShapeType::TextPlainText);

        // Apply the "Bold' and "Italic" formatting settings to the text using the respective properties.
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
        ASSERT_EQ(ShapeType::TextPlainText, shape->get_ShapeType());

        // Use the "On" property to show/hide the text.
        shape = AppendWordArt(doc, u"On set to \"true\"", u"Calibri", 150, 24, System::Drawing::Color::get_Yellow(), System::Drawing::Color::get_Red(),
                              ShapeType::TextPlainText);
        shape->get_TextPath()->set_On(true);

        shape = AppendWordArt(doc, u"On set to \"false\"", u"Calibri", 150, 24, System::Drawing::Color::get_Yellow(), System::Drawing::Color::get_Purple(),
                              ShapeType::TextPlainText);
        shape->get_TextPath()->set_On(false);

        // Use the "Kerning" property to enable/disable kerning spacing between certain characters.
        shape = AppendWordArt(doc, u"Kerning: VAV", u"Times New Roman", 90, 24, System::Drawing::Color::get_Orange(), System::Drawing::Color::get_Red(),
                              ShapeType::TextPlainText);
        shape->get_TextPath()->set_Kerning(true);

        shape = AppendWordArt(doc, u"No kerning: VAV", u"Times New Roman", 100, 24, System::Drawing::Color::get_Orange(), System::Drawing::Color::get_Red(),
                              ShapeType::TextPlainText);
        shape->get_TextPath()->set_Kerning(false);

        // Use the "Spacing" property to set the custom spacing between characters on a scale from 0.0 (none) to 1.0 (default).
        shape = AppendWordArt(doc, u"Spacing set to 0.1", u"Calibri", 120, 24, System::Drawing::Color::get_BlueViolet(), System::Drawing::Color::get_Blue(),
                              ShapeType::TextCascadeDown);
        shape->get_TextPath()->set_Spacing(0.1);

        // Set the "RotateLetters" property to "true" to rotate each character 90 degrees counterclockwise.
        shape = AppendWordArt(doc, u"RotateLetters", u"Calibri", 200, 36, System::Drawing::Color::get_GreenYellow(), System::Drawing::Color::get_Green(),
                              ShapeType::TextWave);
        shape->get_TextPath()->set_RotateLetters(true);

        // Set the "SameLetterHeights" property to "true" to get the x-height of each character to equal the cap height.
        shape = AppendWordArt(doc, u"Same character height for lower and UPPER case", u"Calibri", 300, 24, System::Drawing::Color::get_DeepSkyBlue(),
                              System::Drawing::Color::get_DodgerBlue(), ShapeType::TextSlantUp);
        shape->get_TextPath()->set_SameLetterHeights(true);

        // By default, the text's size will always scale to fit the containing shape's size, overriding the text size setting.
        shape = AppendWordArt(doc, u"FitShape on", u"Calibri", 160, 24, System::Drawing::Color::get_LightBlue(), System::Drawing::Color::get_Blue(),
                              ShapeType::TextPlainText);
        ASSERT_TRUE(shape->get_TextPath()->get_FitShape());
        shape->get_TextPath()->set_Size(24.0);

        // If we set the "FitShape: property to "false", the text will keep the size
        // which the "Size" property specifies regardless of the size of the shape.
        // Use the "TextPathAlignment" property also to align the text to a side of the shape.
        shape = AppendWordArt(doc, u"FitShape off", u"Calibri", 160, 24, System::Drawing::Color::get_LightBlue(), System::Drawing::Color::get_Blue(),
                              ShapeType::TextPlainText);
        shape->get_TextPath()->set_FitShape(false);
        shape->get_TextPath()->set_Size(24.0);
        shape->get_TextPath()->set_TextPathAlignment(TextPathAlignment::Right);

        doc->Save(ArtifactsDir + u"Shape.InsertTextPaths.docx");
        TestInsertTextPaths(ArtifactsDir + u"Shape.InsertTextPaths.docx");
        //ExSkip
    }

    /// <summary>
    /// Insert a new paragraph with a WordArt shape inside it.
    /// </summary>
    static SharedPtr<Shape> AppendWordArt(SharedPtr<Document> doc, String text, String textFontFamily, double shapeWidth, double shapeHeight,
                                          System::Drawing::Color wordArtFill, System::Drawing::Color line, ShapeType wordArtShapeType)
    {
        // Create an inline Shape, which will serve as a container for our WordArt.
        // The shape can only be a valid WordArt shape if we assign a WordArt-designated ShapeType to it.
        // These types will have "WordArt object" in the description,
        // and their enumerator constant names will all start with "Text".
        auto shape = MakeObject<Shape>(doc, wordArtShapeType);
        shape->set_WrapType(WrapType::Inline);
        shape->set_Width(shapeWidth);
        shape->set_Height(shapeHeight);
        shape->set_FillColor(wordArtFill);
        shape->set_StrokeColor(line);

        shape->get_TextPath()->set_Text(text);
        shape->get_TextPath()->set_FontFamily(textFontFamily);

        auto para = System::DynamicCast<Paragraph>(doc->get_FirstSection()->get_Body()->AppendChild(MakeObject<Paragraph>(doc)));
        para->AppendChild(shape);
        return shape;
    }
    //ExEnd

    void TestInsertTextPaths(String filename)
    {
        auto doc = MakeObject<Document>(filename);
        auto shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_ToList();

        TestUtil::VerifyShape(ShapeType::TextPlainText, String::Empty, 480, 24, 0.0, 0.0, shapes->idx_get(0));
        ASSERT_TRUE(shapes->idx_get(0)->get_TextPath()->get_Bold());
        ASSERT_TRUE(shapes->idx_get(0)->get_TextPath()->get_Italic());

        TestUtil::VerifyShape(ShapeType::TextPlainText, String::Empty, 150, 24, 0.0, 0.0, shapes->idx_get(1));
        ASSERT_TRUE(shapes->idx_get(1)->get_TextPath()->get_On());

        TestUtil::VerifyShape(ShapeType::TextPlainText, String::Empty, 150, 24, 0.0, 0.0, shapes->idx_get(2));
        ASSERT_FALSE(shapes->idx_get(2)->get_TextPath()->get_On());

        TestUtil::VerifyShape(ShapeType::TextPlainText, String::Empty, 90, 24, 0.0, 0.0, shapes->idx_get(3));
        ASSERT_TRUE(shapes->idx_get(3)->get_TextPath()->get_Kerning());

        TestUtil::VerifyShape(ShapeType::TextPlainText, String::Empty, 100, 24, 0.0, 0.0, shapes->idx_get(4));
        ASSERT_FALSE(shapes->idx_get(4)->get_TextPath()->get_Kerning());

        TestUtil::VerifyShape(ShapeType::TextCascadeDown, String::Empty, 120, 24, 0.0, 0.0, shapes->idx_get(5));
        ASSERT_NEAR(0.1, shapes->idx_get(5)->get_TextPath()->get_Spacing(), 0.01);

        TestUtil::VerifyShape(ShapeType::TextWave, String::Empty, 200, 36, 0.0, 0.0, shapes->idx_get(6));
        ASSERT_TRUE(shapes->idx_get(6)->get_TextPath()->get_RotateLetters());

        TestUtil::VerifyShape(ShapeType::TextSlantUp, String::Empty, 300, 24, 0.0, 0.0, shapes->idx_get(7));
        ASSERT_TRUE(shapes->idx_get(7)->get_TextPath()->get_SameLetterHeights());

        TestUtil::VerifyShape(ShapeType::TextPlainText, String::Empty, 160, 24, 0.0, 0.0, shapes->idx_get(8));
        ASSERT_TRUE(shapes->idx_get(8)->get_TextPath()->get_FitShape());
        ASPOSE_ASSERT_EQ(24.0, shapes->idx_get(8)->get_TextPath()->get_Size());

        TestUtil::VerifyShape(ShapeType::TextPlainText, String::Empty, 160, 24, 0.0, 0.0, shapes->idx_get(9));
        ASSERT_FALSE(shapes->idx_get(9)->get_TextPath()->get_FitShape());
        ASPOSE_ASSERT_EQ(24.0, shapes->idx_get(9)->get_TextPath()->get_Size());
        ASSERT_EQ(TextPathAlignment::Right, shapes->idx_get(9)->get_TextPath()->get_TextPathAlignment());
    }

    void ShapeRevision()
    {
        //ExStart
        //ExFor:ShapeBase.IsDeleteRevision
        //ExFor:ShapeBase.IsInsertRevision
        //ExSummary:Shows how to work with revision shapes.
        auto doc = MakeObject<Document>();

        ASSERT_FALSE(doc->get_TrackRevisions());

        // Insert an inline shape without tracking revisions, which will make this shape not a revision of any kind.
        auto shape = MakeObject<Shape>(doc, ShapeType::Cube);
        shape->set_WrapType(WrapType::Inline);
        shape->set_Width(100.0);
        shape->set_Height(100.0);
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(shape);

        // Start tracking revisions and then insert another shape, which will be a revision.
        doc->StartTrackRevisions(u"John Doe");

        shape = MakeObject<Shape>(doc, ShapeType::Sun);
        shape->set_WrapType(WrapType::Inline);
        shape->set_Width(100.0);
        shape->set_Height(100.0);
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(shape);

        ArrayPtr<SharedPtr<Shape>> shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_ToArray();

        ASSERT_EQ(2, shapes->get_Length());

        shapes[0]->Remove();

        // Since we removed that shape while we were tracking changes,
        // the shape persists in the document and counts as a delete revision.
        // Accepting this revision will remove the shape permanently, and rejecting it will keep it in the document.
        ASSERT_EQ(ShapeType::Cube, shapes[0]->get_ShapeType());
        ASSERT_TRUE(shapes[0]->get_IsDeleteRevision());

        // And we inserted another shape while tracking changes, so that shape will count as an insert revision.
        // Accepting this revision will assimilate this shape into the document as a non-revision,
        // and rejecting the revision will remove this shape permanently.
        ASSERT_EQ(ShapeType::Sun, shapes[1]->get_ShapeType());
        ASSERT_TRUE(shapes[1]->get_IsInsertRevision());
        //ExEnd
    }

    void MoveRevisions()
    {
        //ExStart
        //ExFor:ShapeBase.IsMoveFromRevision
        //ExFor:ShapeBase.IsMoveToRevision
        //ExSummary:Shows how to identify move revision shapes.
        // A move revision is when we move an element in the document body by cut-and-pasting it in Microsoft Word while
        // tracking changes. If we involve an inline shape in such a text movement, that shape will also be a revision.
        // Copying-and-pasting or moving floating shapes do not create move revisions.
        auto doc = MakeObject<Document>(MyDir + u"Revision shape.docx");

        // Move revisions consist of pairs of "Move from", and "Move to" revisions. We moved in this document in one shape,
        // but until we accept or reject the move revision, there will be two instances of that shape.
        ArrayPtr<SharedPtr<Shape>> shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_ToArray();

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

    void AdjustWithEffects()
    {
        //ExStart
        //ExFor:ShapeBase.AdjustWithEffects(RectangleF)
        //ExFor:ShapeBase.BoundsWithEffects
        //ExSummary:Shows how to check how a shape's bounds are affected by shape effects.
        auto doc = MakeObject<Document>(MyDir + u"Shape shadow effect.docx");

        ArrayPtr<SharedPtr<Shape>> shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_ToArray();

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

        SharedPtr<Shape> shape = shapes[0];

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
        ASPOSE_ASSERT_EQ(1132, rectangleFOut.get_Height());

        // The effects have also affected the visible bounds of the shape.
        ASPOSE_ASSERT_EQ(-28.5, shape->get_BoundsWithEffects().get_X());
        ASPOSE_ASSERT_EQ(-33, shape->get_BoundsWithEffects().get_Y());
        ASPOSE_ASSERT_EQ(192, shape->get_BoundsWithEffects().get_Width());
        ASPOSE_ASSERT_EQ(279, shape->get_BoundsWithEffects().get_Height());
        //ExEnd
    }

    void RenderAllShapes()
    {
        //ExStart
        //ExFor:ShapeBase.GetShapeRenderer
        //ExFor:NodeRendererBase.Save(Stream, ImageSaveOptions)
        //ExSummary:Shows how to use a shape renderer to export shapes to files in the local file system.
        auto doc = MakeObject<Document>(MyDir + u"Various shapes.docx");
        ArrayPtr<SharedPtr<Shape>> shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_ToArray();

        ASSERT_EQ(7, shapes->get_Length());

        // There are 7 shapes in the document, including one group shape with 2 child shapes.
        // We will render every shape to an image file in the local file system
        // while ignoring the group shapes since they have no appearance.
        // This will produce 6 image files.
        for (const auto& shape : System::IterateOver(doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()))
        {
            SharedPtr<ShapeRenderer> renderer = shape->GetShapeRenderer();
            auto options = MakeObject<ImageSaveOptions>(SaveFormat::Png);
            renderer->Save(ArtifactsDir + String::Format(u"Shape.RenderAllShapes.{0}.png", shape->get_Name()), options);
        }
        //ExEnd
    }

    void DocumentHasSmartArtObject()
    {
        //ExStart
        //ExFor:Shape.HasSmartArt
        //ExSummary:Shows how to count the number of shapes in a document with SmartArt objects.
        auto doc = MakeObject<Document>(MyDir + u"SmartArt.docx");

        int numberOfSmartArtShapes = doc->GetChildNodes(NodeType::Shape, true)
                                         ->LINQ_Cast<SharedPtr<Shape>>()
                                         ->LINQ_Count([](SharedPtr<Shape> shape) { return shape->get_HasSmartArt(); });

        ASSERT_EQ(2, numberOfSmartArtShapes);
        //ExEnd
    }

    void OfficeMathRenderer_()
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
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

        auto officeMath = System::DynamicCast<OfficeMath>(doc->GetChild(NodeType::OfficeMath, 0, true));
        auto renderer = MakeObject<OfficeMathRenderer>(officeMath);

        // Verify the size of the image that the OfficeMath object will create when we render it.
        ASSERT_NEAR(119.0f, renderer->get_SizeInPoints().get_Width(), 0.2f);
        ASSERT_NEAR(13.0f, renderer->get_SizeInPoints().get_Height(), 0.1f);

        ASSERT_NEAR(119.0f, renderer->get_BoundsInPoints().get_Width(), 0.2f);
        ASSERT_NEAR(13.0f, renderer->get_BoundsInPoints().get_Height(), 0.1f);

        // Shapes with transparent parts may contain different values in the "OpaqueBoundsInPoints" properties.
        ASSERT_NEAR(119.0f, renderer->get_OpaqueBoundsInPoints().get_Width(), 0.2f);
        ASSERT_NEAR(14.2f, renderer->get_OpaqueBoundsInPoints().get_Height(), 0.1f);

        // Get the shape size in pixels, with linear scaling to a specific DPI.
        System::Drawing::Rectangle bounds = renderer->GetBoundsInPixels(1.0f, 96.0f);

        ASSERT_EQ(159, bounds.get_Width());
        ASSERT_EQ(18, bounds.get_Height());

        // Get the shape size in pixels, but with a different DPI for the horizontal and vertical dimensions.
        bounds = renderer->GetBoundsInPixels(1.0f, 96.0f, 150.0f);
        ASSERT_EQ(159, bounds.get_Width());
        ASSERT_EQ(28, bounds.get_Height());

        // The opaque bounds may vary here also.
        bounds = renderer->GetOpaqueBoundsInPixels(1.0f, 96.0f);

        ASSERT_EQ(159, bounds.get_Width());
        ASSERT_EQ(18, bounds.get_Height());

        bounds = renderer->GetOpaqueBoundsInPixels(1.0f, 96.0f, 150.0f);

        ASSERT_EQ(159, bounds.get_Width());
        ASSERT_EQ(30, bounds.get_Height());
        //ExEnd
    }

    void ShapeTypes()
    {
        //ExStart
        //ExFor:ShapeType
        //ExSummary:Shows how Aspose.Words identify shapes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertShape(ShapeType::Heptagon, RelativeHorizontalPosition::Page, 0, RelativeVerticalPosition::Page, 0, 0, 0, WrapType::None);

        builder->InsertShape(ShapeType::Cloud, RelativeHorizontalPosition::RightMargin, 0, RelativeVerticalPosition::Page, 0, 0, 0, WrapType::None);

        builder->InsertShape(ShapeType::MathPlus, RelativeHorizontalPosition::RightMargin, 0, RelativeVerticalPosition::Page, 0, 0, 0, WrapType::None);

        // To correct identify shape types you need to work with shapes as DML.
        auto saveOptions = MakeObject<OoxmlSaveOptions>(SaveFormat::Docx);
        saveOptions->set_Compliance(OoxmlCompliance::Iso29500_2008_Transitional);

        doc->Save(ArtifactsDir + u"Shape.ShapeTypes.docx", saveOptions);
        doc = MakeObject<Document>(ArtifactsDir + u"Shape.ShapeTypes.docx");

        ArrayPtr<SharedPtr<Shape>> shapes = doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_ToArray();

        for (SharedPtr<Shape> shape : shapes)
        {
            std::cout << System::EnumGetName(shape->get_ShapeType()) << std::endl;
        }

        //ExEnd
    }

    void IsDecorative()
    {
        //ExStart
        //ExFor:ShapeBase.IsDecorative
        //ExSummary:Shows how to set that the shape is decorative.
        auto doc = MakeObject<Document>(MyDir + u"Decorative shapes.docx");

        auto shape = System::DynamicCast<Shape>(doc->GetChildNodes(NodeType::Shape, true)->idx_get(0));
        ASSERT_TRUE(shape->get_IsDecorative());

        // If "AlternativeText" is not empty, the shape cannot be decorative.
        // That's why our value has changed to 'false'.
        shape->set_AlternativeText(u"Alternative text.");
        ASSERT_FALSE(shape->get_IsDecorative());

        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->MoveToDocumentEnd();
        // Create a new shape as decorative.
        shape = builder->InsertShape(ShapeType::Rectangle, 100, 100);
        shape->set_IsDecorative(true);

        doc->Save(ArtifactsDir + u"Shape.IsDecorative.docx");
        //ExEnd
    }
};

} // namespace ApiExamples
