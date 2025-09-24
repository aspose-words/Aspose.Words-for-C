// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDrawing.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/io/stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file_access.h>
#include <system/io/file.h>
#include <system/enum_helpers.h>
#include <system/enum.h>
#include <system/details/dispose_guard.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBox.h>
#include <Aspose.Words.Cpp/Model/Drawing/Stroke.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/LayoutFlow.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageSize.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Drawing/FillType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>

#include "TestUtil.h"
#include "DocumentHelper.h"


using namespace Aspose::Words::Drawing;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2489766318u, ::Aspose::Words::ApiExamples::ExDrawing::ShapeGroupPrinter, ThisTypeBaseTypesInfo);

ExDrawing::ShapeGroupPrinter::ShapeGroupPrinter()
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
}

System::String ExDrawing::ShapeGroupPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDrawing::ShapeGroupPrinter::VisitGroupShapeStart(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape)
{
    mBuilder->AppendLine(u"Shape group started:");
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDrawing::ShapeGroupPrinter::VisitGroupShapeEnd(System::SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape)
{
    mBuilder->AppendLine(u"End of shape group");
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDrawing::ShapeGroupPrinter::VisitShapeStart(System::SharedPtr<Aspose::Words::Drawing::Shape> shape)
{
    mBuilder->AppendLine(System::String(u"\tShape - ") + System::ObjectExt::ToString(shape->get_ShapeType()) + u":");
    mBuilder->AppendLine(System::String(u"\t\tWidth: ") + shape->get_Width());
    mBuilder->AppendLine(System::String(u"\t\tHeight: ") + shape->get_Height());
    mBuilder->AppendLine(System::String(u"\t\tStroke color: ") + shape->get_Stroke()->get_Color());
    mBuilder->AppendLine(System::String(u"\t\tFill color: ") + shape->get_Fill()->get_ForeColor());
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDrawing::ShapeGroupPrinter::VisitShapeEnd(System::SharedPtr<Aspose::Words::Drawing::Shape> shape)
{
    mBuilder->AppendLine(u"\tEnd of shape");
    return Aspose::Words::VisitorAction::Continue;
}


RTTI_INFO_IMPL_HASH(1131775854u, ::Aspose::Words::ApiExamples::ExDrawing, ThisTypeBaseTypesInfo);

void ExDrawing::TestGroupShapes(System::SharedPtr<Aspose::Words::Document> doc)
{
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    auto shapes = System::ExplicitCast<Aspose::Words::Drawing::GroupShape>(doc->GetChild(Aspose::Words::NodeType::GroupShape, 0, true));
    
    ASSERT_EQ(2, shapes->GetChildNodes(Aspose::Words::NodeType::Any, false)->get_Count());
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(shapes->GetChildNodes(Aspose::Words::NodeType::Any, false)->idx_get(0));
    
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::Balloon, shape->get_ShapeType());
    ASPOSE_ASSERT_EQ(200.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(200.0, shape->get_Height());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), shape->get_StrokeColor().ToArgb());
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(shapes->GetChildNodes(Aspose::Words::NodeType::Any, false)->idx_get(1));
    
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::Cube, shape->get_ShapeType());
    ASPOSE_ASSERT_EQ(100.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(100.0, shape->get_Height());
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), shape->get_StrokeColor().ToArgb());
}


namespace gtest_test
{

class ExDrawing : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExDrawing> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExDrawing>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExDrawing> ExDrawing::s_instance;

} // namespace gtest_test

void ExDrawing::TypeOfImage()
{
    //ExStart
    //ExFor:ImageType
    //ExSummary:Shows how to add an image to a shape and check its type.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> imgShape = builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    ASSERT_EQ(Aspose::Words::Drawing::ImageType::Jpeg, imgShape->get_ImageData()->get_ImageType());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDrawing, TypeOfImage)
{
    s_instance->TypeOfImage();
}

} // namespace gtest_test

void ExDrawing::FillSolid()
{
    //ExStart
    //ExFor:Fill.Color
    //ExFor:FillType
    //ExFor:Fill.FillType
    //ExFor:Fill.Solid
    //ExFor:Fill.Transparency
    //ExFor:Font.Fill
    //ExSummary:Shows how to convert any of the fills back to solid fill.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Two color gradient.docx");
    
    // Get Fill object for Font of the first Run.
    System::SharedPtr<Aspose::Words::Drawing::Fill> fill = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0)->get_Font()->get_Fill();
    
    // Check Fill properties of the Font.
    std::cout << System::String::Format(u"The type of the fill is: {0}", fill->get_FillType()) << std::endl;
    std::cout << "The foreground color of the fill is: " << fill->get_ForeColor() << std::endl;
    std::cout << "The fill is transparent at " << (fill->get_Transparency() * 100) << "%" << std::endl;
    
    // Change type of the fill to Solid with uniform green color.
    fill->Solid();
    std::cout << "\nThe fill is changed:" << std::endl;
    std::cout << System::String::Format(u"The type of the fill is: {0}", fill->get_FillType()) << std::endl;
    std::cout << "The foreground color of the fill is: " << fill->get_ForeColor() << std::endl;
    std::cout << "The fill transparency is " << (fill->get_Transparency() * 100) << "%" << std::endl;
    
    doc->Save(get_ArtifactsDir() + u"Drawing.FillSolid.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDrawing, FillSolid)
{
    s_instance->FillSolid();
}

} // namespace gtest_test

void ExDrawing::StrokePattern()
{
    //ExStart
    //ExFor:Stroke.Color2
    //ExFor:Stroke.ImageBytes
    //ExSummary:Shows how to process shape stroke features.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Shape stroke pattern border.docx");
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    System::SharedPtr<Aspose::Words::Drawing::Stroke> stroke = shape->get_Stroke();
    
    // Strokes can have two colors, which are used to create a pattern defined by two-tone image data.
    // Strokes with a single color do not use the Color2 property.
    ASPOSE_ASSERT_EQ(System::Drawing::Color::FromArgb(255, 128, 0, 0), stroke->get_Color());
    ASPOSE_ASSERT_EQ(System::Drawing::Color::FromArgb(255, 255, 255, 0), stroke->get_Color2());
    
    ASSERT_FALSE(System::TestTools::IsNull(stroke->get_ImageBytes()));
    System::IO::File::WriteAllBytes(get_ArtifactsDir() + u"Drawing.StrokePattern.png", stroke->get_ImageBytes());
    //ExEnd
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(8, 8, get_ArtifactsDir() + u"Drawing.StrokePattern.png");
}

namespace gtest_test
{

TEST_F(ExDrawing, StrokePattern)
{
    s_instance->StrokePattern();
}

} // namespace gtest_test

void ExDrawing::GroupOfShapes()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // If you need to create "NonPrimitive" shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
    // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
    // please use DocumentBuilder.InsertShape methods.
    auto balloon = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Balloon);
    balloon->set_Width(200);
    balloon->set_Height(200);
    balloon->get_Stroke()->set_Color(System::Drawing::Color::get_Red());
    
    auto cube = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Cube);
    cube->set_Width(100);
    cube->set_Height(100);
    cube->get_Stroke()->set_Color(System::Drawing::Color::get_Blue());
    
    auto group = System::MakeObject<Aspose::Words::Drawing::GroupShape>(doc);
    group->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(balloon);
    group->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(cube);
    
    ASSERT_TRUE(group->get_IsGroup());
    
    builder->InsertNode(group);
    
    auto printer = System::MakeObject<Aspose::Words::ApiExamples::ExDrawing::ShapeGroupPrinter>();
    group->Accept(printer);
    
    std::cout << printer->GetText() << std::endl;
    TestGroupShapes(doc);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDrawing, GroupOfShapes)
{
    s_instance->GroupOfShapes();
}

} // namespace gtest_test

void ExDrawing::TextBox()
{
    //ExStart
    //ExFor:LayoutFlow
    //ExSummary:Shows how to add text to a text box, and change its orientation
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto textbox = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::TextBox);
    textbox->set_Width(100);
    textbox->set_Height(100);
    textbox->get_TextBox()->set_LayoutFlow(Aspose::Words::Drawing::LayoutFlow::BottomToTop);
    
    textbox->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(System::MakeObject<Aspose::Words::Paragraph>(doc));
    builder->InsertNode(textbox);
    
    builder->MoveTo(textbox->get_FirstParagraph());
    builder->Write(u"This text is flipped 90 degrees to the left.");
    
    doc->Save(get_ArtifactsDir() + u"Drawing.TextBox.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Drawing.TextBox.docx");
    textbox = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::TextBox, textbox->get_ShapeType());
    ASPOSE_ASSERT_EQ(100.0, textbox->get_Width());
    ASPOSE_ASSERT_EQ(100.0, textbox->get_Height());
    ASSERT_EQ(Aspose::Words::Drawing::LayoutFlow::BottomToTop, textbox->get_TextBox()->get_LayoutFlow());
    ASSERT_EQ(u"This text is flipped 90 degrees to the left.", textbox->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDrawing, TextBox)
{
    s_instance->TextBox();
}

} // namespace gtest_test

void ExDrawing::GetDataFromImage()
{
    //ExStart
    //ExFor:ImageData.ImageBytes
    //ExFor:ImageData.ToByteArray
    //ExFor:ImageData.ToStream
    //ExSummary:Shows how to create an image file from a shape's raw image data.
    auto imgSourceDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Images.docx");
    ASSERT_EQ(10, imgSourceDoc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    //ExSkip
    
    auto imgShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(imgSourceDoc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_TRUE(imgShape->get_HasImage());
    
    // ToByteArray() returns the array stored in the ImageBytes property.
    ASPOSE_ASSERT_EQ(imgShape->get_ImageData()->get_ImageBytes(), imgShape->get_ImageData()->ToByteArray());
    
    // Save the shape's image data to an image file in the local file system.
    {
        System::SharedPtr<System::IO::Stream> imgStream = imgShape->get_ImageData()->ToStream();
        {
            auto outStream = System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + u"Drawing.GetDataFromImage.png", System::IO::FileMode::Create, System::IO::FileAccess::ReadWrite);
            imgStream->CopyTo(outStream);
        }
    }
    //ExEnd
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(2467, 1500, get_ArtifactsDir() + u"Drawing.GetDataFromImage.png");
}

namespace gtest_test
{

TEST_F(ExDrawing, GetDataFromImage)
{
    s_instance->GetDataFromImage();
}

} // namespace gtest_test

void ExDrawing::ImageData()
{
    //ExStart
    //ExFor:ImageData.BiLevel
    //ExFor:ImageData.Borders
    //ExFor:ImageData.Brightness
    //ExFor:ImageData.ChromaKey
    //ExFor:ImageData.Contrast
    //ExFor:ImageData.CropBottom
    //ExFor:ImageData.CropLeft
    //ExFor:ImageData.CropRight
    //ExFor:ImageData.CropTop
    //ExFor:ImageData.GrayScale
    //ExFor:ImageData.IsLink
    //ExFor:ImageData.IsLinkOnly
    //ExFor:ImageData.Title
    //ExSummary:Shows how to edit a shape's image data.
    auto imgSourceDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Images.docx");
    auto sourceShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(imgSourceDoc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->idx_get(0));
    
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    
    // Import a shape from the source document and append it to the first paragraph.
    auto importedShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(dstDoc->ImportNode(sourceShape, true));
    dstDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(importedShape);
    
    // The imported shape contains an image. We can access the image's properties and raw data via the ImageData object.
    System::SharedPtr<Aspose::Words::Drawing::ImageData> imageData = importedShape->get_ImageData();
    imageData->set_Title(u"Imported Image");
    
    ASSERT_TRUE(imageData->get_HasImage());
    
    // If an image has no borders, its ImageData object will define the border color as empty.
    ASSERT_EQ(4, imageData->get_Borders()->get_Count());
    ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty, imageData->get_Borders()->idx_get(0)->get_Color());
    
    // This image does not link to another shape or image file in the local file system.
    ASSERT_FALSE(imageData->get_IsLink());
    ASSERT_FALSE(imageData->get_IsLinkOnly());
    
    // The "Brightness" and "Contrast" properties define image brightness and contrast
    // on a 0-1 scale, with the default value at 0.5.
    imageData->set_Brightness(0.8);
    imageData->set_Contrast(1.0);
    
    // The above brightness and contrast values have created an image with a lot of white.
    // We can select a color with the ChromaKey property to replace with transparency, such as white.
    imageData->set_ChromaKey(System::Drawing::Color::get_White());
    
    // Import the source shape again and set the image to monochrome.
    importedShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(dstDoc->ImportNode(sourceShape, true));
    dstDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(importedShape);
    
    importedShape->get_ImageData()->set_GrayScale(true);
    
    // Import the source shape again to create a third image and set it to BiLevel.
    // BiLevel sets every pixel to either black or white, whichever is closer to the original color.
    importedShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(dstDoc->ImportNode(sourceShape, true));
    dstDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(importedShape);
    
    importedShape->get_ImageData()->set_BiLevel(true);
    
    // Cropping is determined on a 0-1 scale. Cropping a side by 0.3
    // will crop 30% of the image out at the cropped side.
    importedShape->get_ImageData()->set_CropBottom(0.3);
    importedShape->get_ImageData()->set_CropLeft(0.3);
    importedShape->get_ImageData()->set_CropTop(0.3);
    importedShape->get_ImageData()->set_CropRight(0.3);
    
    dstDoc->Save(get_ArtifactsDir() + u"Drawing.ImageData.docx");
    //ExEnd
    
    imgSourceDoc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Drawing.ImageData.docx");
    sourceShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(imgSourceDoc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(2467, 1500, Aspose::Words::Drawing::ImageType::Jpeg, sourceShape);
    ASSERT_EQ(u"Imported Image", sourceShape->get_ImageData()->get_Title());
    ASSERT_NEAR(0.8, sourceShape->get_ImageData()->get_Brightness(), 0.1);
    ASSERT_NEAR(1.0, sourceShape->get_ImageData()->get_Contrast(), 0.1);
    ASSERT_EQ(System::Drawing::Color::get_White().ToArgb(), sourceShape->get_ImageData()->get_ChromaKey().ToArgb());
    
    sourceShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(imgSourceDoc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(2467, 1500, Aspose::Words::Drawing::ImageType::Jpeg, sourceShape);
    ASSERT_TRUE(sourceShape->get_ImageData()->get_GrayScale());
    
    sourceShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(imgSourceDoc->GetChild(Aspose::Words::NodeType::Shape, 2, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(2467, 1500, Aspose::Words::Drawing::ImageType::Jpeg, sourceShape);
    ASSERT_TRUE(sourceShape->get_ImageData()->get_BiLevel());
    ASSERT_NEAR(0.3, sourceShape->get_ImageData()->get_CropBottom(), 0.1);
    ASSERT_NEAR(0.3, sourceShape->get_ImageData()->get_CropLeft(), 0.1);
    ASSERT_NEAR(0.3, sourceShape->get_ImageData()->get_CropTop(), 0.1);
    ASSERT_NEAR(0.3, sourceShape->get_ImageData()->get_CropRight(), 0.1);
}

namespace gtest_test
{

TEST_F(ExDrawing, ImageData)
{
    s_instance->ImageData();
}

} // namespace gtest_test

void ExDrawing::ImageSize()
{
    //ExStart
    //ExFor:ImageSize.HeightPixels
    //ExFor:ImageSize.HorizontalResolution
    //ExFor:ImageSize.VerticalResolution
    //ExFor:ImageSize.WidthPixels
    //ExSummary:Shows how to read the properties of an image in a shape.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a shape into the document which contains an image taken from our local file system.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    // If the shape contains an image, its ImageData property will be valid,
    // and it will contain an ImageSize object.
    System::SharedPtr<Aspose::Words::Drawing::ImageSize> imageSize = shape->get_ImageData()->get_ImageSize();
    
    // The ImageSize object contains read-only information about the image within the shape.
    ASSERT_EQ(400, imageSize->get_HeightPixels());
    ASSERT_EQ(400, imageSize->get_WidthPixels());
    
    const double delta = 0.05;
    ASSERT_NEAR(95.98, imageSize->get_HorizontalResolution(), delta);
    ASSERT_NEAR(95.98, imageSize->get_VerticalResolution(), delta);
    
    // We can base the size of the shape on the size of its image to avoid stretching the image.
    shape->set_Width(imageSize->get_WidthPoints() * 2);
    shape->set_Height(imageSize->get_HeightPoints() * 2);
    
    doc->Save(get_ArtifactsDir() + u"Drawing.ImageSize.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Drawing.ImageSize.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, shape);
    ASPOSE_ASSERT_EQ(600.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(600.0, shape->get_Height());
    
    imageSize = shape->get_ImageData()->get_ImageSize();
    
    ASSERT_EQ(400, imageSize->get_HeightPixels());
    ASSERT_EQ(400, imageSize->get_WidthPixels());
    ASSERT_NEAR(95.98, imageSize->get_HorizontalResolution(), delta);
    ASSERT_NEAR(95.98, imageSize->get_VerticalResolution(), delta);
}

namespace gtest_test
{

TEST_F(ExDrawing, ImageSize)
{
    s_instance->ImageSize();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
