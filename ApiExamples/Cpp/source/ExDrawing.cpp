// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDrawing.h"

#include <testing/test_predicates.h>
#include <system/text/string_builder.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/memory_stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file_info.h>
#include <system/io/file_access.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enum_helpers.h>
#include <system/details/dispose_guard.h>
#include <system/console.h>
#include <system/collections/list.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <net/web_client.h>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/rotate_flip_type.h>
#include <drawing/imaging/image_format.h>
#include <drawing/image_format_converter.h>
#include <drawing/image.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBox.h>
#include <Aspose.Words.Cpp/Model/Drawing/Stroke.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/LayoutFlow.h>
#include <Aspose.Words.Cpp/Model/Drawing/JoinStyle.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageSize.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Model/Drawing/FlipOrientation.h>
#include <Aspose.Words.Cpp/Model/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Model/Drawing/EndCap.h>
#include <Aspose.Words.Cpp/Model/Drawing/DashStyle.h>
#include <Aspose.Words.Cpp/Model/Drawing/ArrowWidth.h>
#include <Aspose.Words.Cpp/Model/Drawing/ArrowType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ArrowLength.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>

#include "TestUtil.h"
#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2676422227u, ::ApiExamples::ExDrawing::ShapeInfoPrinter, ThisTypeBaseTypesInfo);

ExDrawing::ShapeInfoPrinter::ShapeInfoPrinter()
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
}

String ExDrawing::ShapeInfoPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDrawing::ShapeInfoPrinter::VisitGroupShapeStart(SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape)
{
    mBuilder->AppendLine(u"Shape group started:");
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDrawing::ShapeInfoPrinter::VisitGroupShapeEnd(SharedPtr<Aspose::Words::Drawing::GroupShape> groupShape)
{
    mBuilder->AppendLine(u"End of shape group");
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDrawing::ShapeInfoPrinter::VisitShapeStart(SharedPtr<Aspose::Words::Drawing::Shape> shape)
{
    mBuilder->AppendLine(String(u"\tShape - ") + System::ObjectExt::ToString(shape->get_ShapeType()) + u":");
    mBuilder->AppendLine(String(u"\t\tWidth: ") + shape->get_Width());
    mBuilder->AppendLine(String(u"\t\tHeight: ") + shape->get_Height());
    mBuilder->AppendLine(String(u"\t\tStroke color: ") + shape->get_Stroke()->get_Color());
    mBuilder->AppendLine(String(u"\t\tFill color: ") + shape->get_Fill()->get_Color());
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDrawing::ShapeInfoPrinter::VisitShapeEnd(SharedPtr<Aspose::Words::Drawing::Shape> shape)
{
    mBuilder->AppendLine(u"\tEnd of shape");
    return Aspose::Words::VisitorAction::Continue;
}

System::Object::shared_members_type ApiExamples::ExDrawing::ShapeInfoPrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExDrawing::ShapeInfoPrinter::mBuilder", this->mBuilder);

    return result;
}

void ExDrawing::TestGroupShapes(SharedPtr<Aspose::Words::Document> doc)
{
    doc = DocumentHelper::SaveOpen(doc);
    auto shapes = System::DynamicCast<Aspose::Words::Drawing::GroupShape>(doc->GetChild(Aspose::Words::NodeType::GroupShape, 0, true));

    ASSERT_EQ(2, shapes->get_ChildNodes()->get_Count());

    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(shapes->get_ChildNodes()->idx_get(0));

    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::Balloon, shape->get_ShapeType());
    ASPOSE_ASSERT_EQ(200.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(200.0, shape->get_Height());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), shape->get_StrokeColor().ToArgb());

    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(shapes->get_ChildNodes()->idx_get(1));

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
    static SharedPtr<::ApiExamples::ExDrawing> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExDrawing>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExDrawing> ExDrawing::s_instance;

} // namespace gtest_test

void ExDrawing::VariousShapes()
{
    //ExStart
    //ExFor:Drawing.ArrowLength
    //ExFor:Drawing.ArrowType
    //ExFor:Drawing.ArrowWidth
    //ExFor:Drawing.DashStyle
    //ExFor:Drawing.EndCap
    //ExFor:Drawing.Fill.Color
    //ExFor:Drawing.Fill.ImageBytes
    //ExFor:Drawing.Fill.On
    //ExFor:Drawing.JoinStyle
    //ExFor:Shape.Stroke
    //ExFor:Stroke.Color
    //ExFor:Stroke.StartArrowLength
    //ExFor:Stroke.StartArrowType
    //ExFor:Stroke.StartArrowWidth
    //ExFor:Stroke.EndArrowLength
    //ExFor:Stroke.EndArrowWidth
    //ExFor:Stroke.DashStyle
    //ExFor:Stroke.EndArrowType
    //ExFor:Stroke.EndCap
    //ExFor:Stroke.Opacity
    //ExSummary:Shows to create a variety of shapes.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Draw a dotted horizontal half-transparent red line with an arrow on the left end and a diamond on the other
    auto arrow = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Line);
    arrow->set_Width(200);
    arrow->get_Stroke()->set_Color(System::Drawing::Color::get_Red());
    arrow->get_Stroke()->set_StartArrowType(Aspose::Words::Drawing::ArrowType::Arrow);
    arrow->get_Stroke()->set_StartArrowLength(Aspose::Words::Drawing::ArrowLength::Long);
    arrow->get_Stroke()->set_StartArrowWidth(Aspose::Words::Drawing::ArrowWidth::Wide);
    arrow->get_Stroke()->set_EndArrowType(Aspose::Words::Drawing::ArrowType::Diamond);
    arrow->get_Stroke()->set_EndArrowLength(Aspose::Words::Drawing::ArrowLength::Long);
    arrow->get_Stroke()->set_EndArrowWidth(Aspose::Words::Drawing::ArrowWidth::Wide);
    arrow->get_Stroke()->set_DashStyle(Aspose::Words::Drawing::DashStyle::Dash);
    arrow->get_Stroke()->set_Opacity(0.5);

    ASSERT_EQ(Aspose::Words::Drawing::JoinStyle::Miter, arrow->get_Stroke()->get_JoinStyle());

    builder->InsertNode(arrow);

    // Draw a thick black diagonal line with rounded ends
    auto line = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Line);
    line->set_Top(40);
    line->set_Width(200);
    line->set_Height(20);
    line->set_StrokeWeight(5.0);
    line->get_Stroke()->set_EndCap(Aspose::Words::Drawing::EndCap::Round);

    builder->InsertNode(line);

    // Draw an arrow with a green fill
    auto filledInArrow = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Arrow);
    filledInArrow->set_Width(200);
    filledInArrow->set_Height(40);
    filledInArrow->set_Top(100);
    filledInArrow->get_Fill()->set_Color(System::Drawing::Color::get_Green());
    filledInArrow->get_Fill()->set_On(true);

    builder->InsertNode(filledInArrow);

    // Draw an arrow filled in with the Aspose logo and flip its orientation
    auto filledInArrowImg = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Arrow);
    filledInArrowImg->set_Width(200);
    filledInArrowImg->set_Height(40);
    filledInArrowImg->set_Top(160);
    filledInArrowImg->set_FlipOrientation(Aspose::Words::Drawing::FlipOrientation::Both);

    {
        auto webClient = MakeObject<System::Net::WebClient>();
        ArrayPtr<uint8_t> imageBytes = webClient->DownloadData(AsposeLogoUrl);

        {
            auto stream = MakeObject<System::IO::MemoryStream>(imageBytes);
            SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromStream(stream);
            // When we flipped the orientation of our arrow, the image content was flipped too
            // If we want it to be displayed the right side up, we have to reverse the arrow flip on the image
            image->RotateFlip(System::Drawing::RotateFlipType::RotateNoneFlipXY);

            filledInArrowImg->get_ImageData()->SetImage(image);
            filledInArrowImg->get_Stroke()->set_JoinStyle(Aspose::Words::Drawing::JoinStyle::Round);

            builder->InsertNode(filledInArrowImg);
        }
    }

    doc->Save(ArtifactsDir + u"Drawing.VariousShapes.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Drawing.VariousShapes.docx");

    ASSERT_EQ(4, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());

    arrow = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::Line, arrow->get_ShapeType());
    ASPOSE_ASSERT_EQ(200.0, arrow->get_Width());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), arrow->get_Stroke()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::Drawing::ArrowType::Arrow, arrow->get_Stroke()->get_StartArrowType());
    ASSERT_EQ(Aspose::Words::Drawing::ArrowLength::Long, arrow->get_Stroke()->get_StartArrowLength());
    ASSERT_EQ(Aspose::Words::Drawing::ArrowWidth::Wide, arrow->get_Stroke()->get_StartArrowWidth());
    ASSERT_EQ(Aspose::Words::Drawing::ArrowType::Diamond, arrow->get_Stroke()->get_EndArrowType());
    ASSERT_EQ(Aspose::Words::Drawing::ArrowLength::Long, arrow->get_Stroke()->get_EndArrowLength());
    ASSERT_EQ(Aspose::Words::Drawing::ArrowWidth::Wide, arrow->get_Stroke()->get_EndArrowWidth());
    ASSERT_EQ(Aspose::Words::Drawing::DashStyle::Dash, arrow->get_Stroke()->get_DashStyle());
    ASPOSE_ASSERT_EQ(0.5, arrow->get_Stroke()->get_Opacity());

    line = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));

    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::Line, line->get_ShapeType());
    ASPOSE_ASSERT_EQ(40.0, line->get_Top());
    ASPOSE_ASSERT_EQ(200.0, line->get_Width());
    ASPOSE_ASSERT_EQ(20.0, line->get_Height());
    ASPOSE_ASSERT_EQ(5.0, line->get_StrokeWeight());
    ASSERT_EQ(Aspose::Words::Drawing::EndCap::Round, line->get_Stroke()->get_EndCap());

    filledInArrow = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));

    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::Arrow, filledInArrow->get_ShapeType());
    ASPOSE_ASSERT_EQ(200.0, filledInArrow->get_Width());
    ASPOSE_ASSERT_EQ(40.0, filledInArrow->get_Height());
    ASPOSE_ASSERT_EQ(100.0, filledInArrow->get_Top());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), filledInArrow->get_Fill()->get_Color().ToArgb());
    ASSERT_TRUE(filledInArrow->get_Fill()->get_On());

    filledInArrowImg = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 3, true));

    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::Arrow, filledInArrowImg->get_ShapeType());
    ASPOSE_ASSERT_EQ(200.0, filledInArrowImg->get_Width());
    ASPOSE_ASSERT_EQ(40.0, filledInArrowImg->get_Height());
    ASPOSE_ASSERT_EQ(160.0, filledInArrowImg->get_Top());
    ASSERT_EQ(Aspose::Words::Drawing::FlipOrientation::Both, filledInArrowImg->get_FlipOrientation());
}

namespace gtest_test
{

TEST_F(ExDrawing, VariousShapes)
{
    s_instance->VariousShapes();
}

} // namespace gtest_test

void ExDrawing::TypeOfImage()
{
    //ExStart
    //ExFor:Drawing.ImageType
    //ExSummary:Shows how to add an image to a shape and check its type.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    {
        auto webClient = MakeObject<System::Net::WebClient>();
        ArrayPtr<uint8_t> imageBytes = webClient->DownloadData(AsposeLogoUrl);

        {
            auto stream = MakeObject<System::IO::MemoryStream>(imageBytes);
            SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromStream(stream);

            // The image started off as an animated .gif but it gets converted to a .png since there cannot be animated images in documents
            SharedPtr<Shape> imgShape = builder->InsertImage(image);
            ASSERT_EQ(Aspose::Words::Drawing::ImageType::Png, imgShape->get_ImageData()->get_ImageType());
        }
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDrawing, TypeOfImage)
{
    s_instance->TypeOfImage();
}

} // namespace gtest_test

void ExDrawing::SaveAllImages()
{
    //ExStart
    //ExFor:ImageData.HasImage
    //ExFor:ImageData.ToImage
    //ExFor:ImageData.Save(Stream)
    //ExSummary:Shows how to save all the images from a document to the file system.
    auto imgSourceDoc = MakeObject<Document>(MyDir + u"Images.docx");

    // Images are stored inside shapes, and if a shape contains an image, its "HasImage" flag will be set
    // Get an enumerator for all shapes with that flag in the document

    auto shapes = imgSourceDoc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape>>()->LINQ_Where([](SharedPtr<Shape> s) { return s->get_HasImage(); });

    // We will use an ImageFormatConverter to determine an image's file extension
    auto formatConverter = MakeObject<System::Drawing::ImageFormatConverter>();

    // Go over all of the document's shapes
    // If a shape contains image data, save the image in the local file system
    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Shape>>> enumerator = shapes->GetEnumerator();
        int shapeIndex = 0;

        while (enumerator->MoveNext())
        {
            SharedPtr<Aspose::Words::Drawing::ImageData> imageData = enumerator->get_Current()->get_ImageData();
            SharedPtr<System::Drawing::Imaging::ImageFormat> format = imageData->ToImage()->get_RawFormat();
            String fileExtension = formatConverter->ConvertToString(format);

            {
                SharedPtr<System::IO::FileStream> fileStream = System::IO::File::Create(ArtifactsDir + String::Format(u"Drawing.SaveAllImages.{0}.{1}",++shapeIndex,fileExtension));
                imageData->Save(fileStream);
            }
        }
    }
    //ExEnd

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(String s)> _local_func_1 = [](String s)
    {
        return s.StartsWith(ArtifactsDir + u"Drawing.SaveAllImages.");
    };

    ArrayPtr<String> imageFileNames = System::IO::Directory::GetFiles(ArtifactsDir)->LINQ_Where(static_cast<System::Func<String, bool>>(_local_func_1))->LINQ_ToArray();

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<SharedPtr<System::IO::FileInfo>(String s)> _local_func_2 = [](String s)
    {
        return MakeObject<System::IO::FileInfo>(s);
    };

    SharedPtr<System::Collections::Generic::List<SharedPtr<System::IO::FileInfo>>> fileInfos = imageFileNames->LINQ_Select(static_cast<System::Func<String, SharedPtr<System::IO::FileInfo>>>(_local_func_2))->LINQ_ToList();

    TestUtil::VerifyImage(2467, 1500, fileInfos->idx_get(0)->get_FullName());
    ASSERT_EQ(u".Jpeg", fileInfos->idx_get(0)->get_Extension());
    TestUtil::VerifyImage(400, 400, fileInfos->idx_get(1)->get_FullName());
    ASSERT_EQ(u".Png", fileInfos->idx_get(1)->get_Extension());
    TestUtil::VerifyImage(382, 138, fileInfos->idx_get(2)->get_FullName());
    ASSERT_EQ(u".Emf", fileInfos->idx_get(2)->get_Extension());
    TestUtil::VerifyImage(1600, 1600, fileInfos->idx_get(3)->get_FullName());
    ASSERT_EQ(u".Wmf", fileInfos->idx_get(3)->get_Extension());
    TestUtil::VerifyImage(534, 534, fileInfos->idx_get(4)->get_FullName());
    ASSERT_EQ(u".Emf", fileInfos->idx_get(4)->get_Extension());
    TestUtil::VerifyImage(1260, 660, fileInfos->idx_get(5)->get_FullName());
    ASSERT_EQ(u".Jpeg", fileInfos->idx_get(5)->get_Extension());
    TestUtil::VerifyImage(1125, 1500, fileInfos->idx_get(6)->get_FullName());
    ASSERT_EQ(u".Jpeg", fileInfos->idx_get(6)->get_Extension());
    TestUtil::VerifyImage(1027, 1500, fileInfos->idx_get(7)->get_FullName());
    ASSERT_EQ(u".Jpeg", fileInfos->idx_get(7)->get_Extension());
    TestUtil::VerifyImage(1200, 1500, fileInfos->idx_get(8)->get_FullName());
    ASSERT_EQ(u".Jpeg", fileInfos->idx_get(8)->get_Extension());
}

namespace gtest_test
{

TEST_F(ExDrawing, SaveAllImages)
{
    s_instance->SaveAllImages();
}

} // namespace gtest_test

void ExDrawing::ImportImage()
{
    //ExStart
    //ExFor:ImageData.SetImage(Image)
    //ExFor:ImageData.SetImage(Stream)
    //ExSummary:Shows two ways of importing images from the local file system into a document.
    auto doc = MakeObject<Document>();

    // We can get an image from a file, set it as the image of a shape and append it to a paragraph
    SharedPtr<System::Drawing::Image> srcImage = System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg");

    auto imgShape = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Image);
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(imgShape);
    imgShape->get_ImageData()->SetImage(srcImage);
    srcImage->Dispose();

    // We can also open an image file using a stream and set its contents as a shape's image
    {
        SharedPtr<System::IO::Stream> stream = MakeObject<System::IO::FileStream>(ImageDir + u"Logo.jpg", System::IO::FileMode::Open, System::IO::FileAccess::Read);
        imgShape = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Image);
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(imgShape);
        imgShape->get_ImageData()->SetImage(stream);
        imgShape->set_Left(150.0f);
    }

    doc->Save(ArtifactsDir + u"Drawing.ImportImage.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Drawing.ImportImage.docx");

    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());

    imgShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imgShape);
    ASPOSE_ASSERT_EQ(0.0, imgShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imgShape->get_Top());
    ASPOSE_ASSERT_EQ(300.0, imgShape->get_Height());
    ASPOSE_ASSERT_EQ(300.0, imgShape->get_Width());
    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imgShape);

    imgShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imgShape);
    ASPOSE_ASSERT_EQ(150.0, imgShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imgShape->get_Top());
    ASPOSE_ASSERT_EQ(300.0, imgShape->get_Height());
    ASPOSE_ASSERT_EQ(300.0, imgShape->get_Width());
}

namespace gtest_test
{

TEST_F(ExDrawing, ImportImage)
{
    s_instance->ImportImage();
}

} // namespace gtest_test

void ExDrawing::StrokePattern()
{
    //ExStart
    //ExFor:Stroke.Color2
    //ExFor:Stroke.ImageBytes
    //ExSummary:Shows how to process shape stroke features.
    // Open a document which contains a rectangle with a thick, two-tone-patterned outline
    auto doc = MakeObject<Document>(MyDir + u"Shape stroke pattern border.docx");

    // Get the first shape's stroke
    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    SharedPtr<Stroke> s = shape->get_Stroke();

    // Every stroke will have a Color attribute, but only strokes from older Word versions have a Color2 attribute,
    // since the two-tone pattern line feature which requires the Color2 attribute is no longer supported
    ASPOSE_ASSERT_EQ(System::Drawing::Color::FromArgb(255, 128, 0, 0), s->get_Color());
    ASPOSE_ASSERT_EQ(System::Drawing::Color::FromArgb(255, 255, 255, 0), s->get_Color2());

    // This attribute contains the image data for the pattern, which we can save to our local file system
    ASSERT_FALSE(s->get_ImageBytes() == nullptr);
    System::IO::File::WriteAllBytes(ArtifactsDir + u"Drawing.StrokePattern.png", s->get_ImageBytes());
    //ExEnd

    TestUtil::VerifyImage(8, 8, ArtifactsDir + u"Drawing.StrokePattern.png");
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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // If you need to create "NonPrimitive" shapes, like SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
    // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
    // please use DocumentBuilder.InsertShape methods
    auto balloon = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Balloon);
    balloon->set_Width(200);
    balloon->set_Height(200);
    balloon->get_Stroke()->set_Color(System::Drawing::Color::get_Red());

    auto cube = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Cube);
    cube->set_Width(100);
    cube->set_Height(100);
    cube->get_Stroke()->set_Color(System::Drawing::Color::get_Blue());

    auto group = MakeObject<GroupShape>(doc);
    group->AppendChild(balloon);
    group->AppendChild(cube);

    ASSERT_TRUE(group->get_IsGroup());

    builder->InsertNode(group);

    auto printer = MakeObject<ExDrawing::ShapeInfoPrinter>();
    group->Accept(printer);

    System::Console::WriteLine(printer->GetText());
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
    //ExFor:Drawing.LayoutFlow
    //ExSummary:Shows how to add text to a textbox and change its orientation
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    auto textbox = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::TextBox);
    textbox->set_Width(100);
    textbox->set_Height(100);
    textbox->get_TextBox()->set_LayoutFlow(Aspose::Words::Drawing::LayoutFlow::BottomToTop);

    textbox->AppendChild(MakeObject<Paragraph>(doc));
    builder->InsertNode(textbox);

    builder->MoveTo(textbox->get_FirstParagraph());
    builder->Write(u"This text is flipped 90 degrees to the left.");

    doc->Save(ArtifactsDir + u"Drawing.TextBox.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Drawing.TextBox.docx");
    textbox = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

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
    //ExSummary:Shows how to access raw image data in a shape's ImageData object.
    auto imgSourceDoc = MakeObject<Document>(MyDir + u"Images.docx");
    ASSERT_EQ(10, imgSourceDoc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    //ExSkip

    // Get a shape from the document that contains an image
    auto imgShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(imgSourceDoc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    // ToByteArray() returns the value of the ImageBytes property
    ASPOSE_ASSERT_EQ(imgShape->get_ImageData()->get_ImageBytes(), imgShape->get_ImageData()->ToByteArray());

    // Put the shape's image data into a stream
    // Then, put the image data from that stream into another stream which creates an image file in the local file system
    {
        SharedPtr<System::IO::Stream> imgStream = imgShape->get_ImageData()->ToStream();
        {
            auto outStream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"Drawing.GetDataFromImage.png", System::IO::FileMode::Create);
            imgStream->CopyTo(outStream);
        }
    }
    //ExEnd

    TestUtil::VerifyImage(2467, 1500, ArtifactsDir + u"Drawing.GetDataFromImage.png");
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
    //ExSummary:Shows how to edit images using the ImageData attribute.
    // Open a document that contains images
    auto imgSourceDoc = MakeObject<Document>(MyDir + u"Images.docx");

    auto sourceShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(imgSourceDoc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->idx_get(0));
    auto dstDoc = MakeObject<Document>();

    // Import a shape from the source document and append it to the first paragraph, effectively cloning it
    auto importedShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(dstDoc->ImportNode(sourceShape, true));
    dstDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(importedShape);

    // Get the ImageData of the imported shape
    SharedPtr<Aspose::Words::Drawing::ImageData> imageData = importedShape->get_ImageData();
    imageData->set_Title(u"Imported Image");

    // If an image appears to have no borders, its ImageData object will still have them, but in an unspecified color
    ASSERT_EQ(4, imageData->get_Borders()->get_Count());
    ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty, imageData->get_Borders()->idx_get(0)->get_Color());

    ASSERT_TRUE(imageData->get_HasImage());

    // This image is not linked to a shape or to an image in the file system
    ASSERT_FALSE(imageData->get_IsLink());
    ASSERT_FALSE(imageData->get_IsLinkOnly());

    // Brightness and contrast are defined on a 0-1 scale, with 0.5 being the default value
    imageData->set_Brightness(0.8);
    imageData->set_Contrast(1.0);

    // Our image will have a lot of white now that we've changed the brightness and contrast like that
    // We can treat white as transparent with the following attribute
    imageData->set_ChromaKey(System::Drawing::Color::get_White());

    // Import the source shape again, set it to black and white
    importedShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(dstDoc->ImportNode(sourceShape, true));
    dstDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(importedShape);

    importedShape->get_ImageData()->set_GrayScale(true);

    // Import the source shape again to create a third image, and set it to BiLevel
    // Unlike greyscale, which preserves the brightness of the original colors,
    // BiLevel sets every pixel to either black or white, whichever is closer to the original color
    importedShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(dstDoc->ImportNode(sourceShape, true));
    dstDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(importedShape);

    importedShape->get_ImageData()->set_BiLevel(true);

    // Cropping is determined on a 0-1 scale
    // Cropping a side by 0.3 will crop 30% of the image out at that side
    importedShape->get_ImageData()->set_CropBottom(0.3);
    importedShape->get_ImageData()->set_CropLeft(0.3);
    importedShape->get_ImageData()->set_CropTop(0.3);
    importedShape->get_ImageData()->set_CropRight(0.3);

    dstDoc->Save(ArtifactsDir + u"Drawing.ImageData.docx");
    //ExEnd

    imgSourceDoc = MakeObject<Document>(ArtifactsDir + u"Drawing.ImageData.docx");
    sourceShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(imgSourceDoc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyImageInShape(2467, 1500, Aspose::Words::Drawing::ImageType::Jpeg, sourceShape);
    ASSERT_EQ(u"Imported Image", sourceShape->get_ImageData()->get_Title());
    ASSERT_NEAR(0.8, sourceShape->get_ImageData()->get_Brightness(), 0.1);
    ASSERT_NEAR(1.0, sourceShape->get_ImageData()->get_Contrast(), 0.1);
    ASSERT_EQ(System::Drawing::Color::get_White().ToArgb(), sourceShape->get_ImageData()->get_ChromaKey().ToArgb());

    sourceShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(imgSourceDoc->GetChild(Aspose::Words::NodeType::Shape, 1, true));

    TestUtil::VerifyImageInShape(2467, 1500, Aspose::Words::Drawing::ImageType::Jpeg, sourceShape);
    ASSERT_TRUE(sourceShape->get_ImageData()->get_GrayScale());

    sourceShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(imgSourceDoc->GetChild(Aspose::Words::NodeType::Shape, 2, true));

    TestUtil::VerifyImageInShape(2467, 1500, Aspose::Words::Drawing::ImageType::Jpeg, sourceShape);
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
    //ExSummary:Shows how to access and use a shape's ImageSize property.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a shape into the document which contains an image taken from our local file system
    SharedPtr<Shape> shape = builder->InsertImage(ImageDir + u"Logo.jpg");

    // If the shape contains an image, its ImageData property will be valid, and it will contain an ImageSize object
    SharedPtr<Aspose::Words::Drawing::ImageSize> imageSize = shape->get_ImageData()->get_ImageSize();

    // The ImageSize object contains raw information about the image within the shape
    ASSERT_EQ(400, imageSize->get_HeightPixels());
    ASSERT_EQ(400, imageSize->get_WidthPixels());

    const double delta = 0.05;
    ASSERT_NEAR(95.98, imageSize->get_HorizontalResolution(), delta);
    ASSERT_NEAR(95.98, imageSize->get_VerticalResolution(), delta);

    // These values are read-only
    // If we want to transform the image, we need to change the size of the shape that contains it
    // We can still use values within ImageSize as a reference
    // In the example below, we will get the shape to display the image in twice its original size
    shape->set_Width(imageSize->get_WidthPoints() * 2);
    shape->set_Height(imageSize->get_HeightPoints() * 2);

    doc->Save(ArtifactsDir + u"Drawing.ImageSize.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Drawing.ImageSize.docx");
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, shape);
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
