// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExImage.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/setters_wrap.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_info.h>
#include <system/io/file.h>
#include <system/func.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/collections/ienumerable.h>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageSize.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "TestUtil.h"


using namespace Aspose::Words::Drawing;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(631314891u, ::Aspose::Words::ApiExamples::ExImage, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExImage : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExImage> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExImage>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExImage> ExImage::s_instance;

} // namespace gtest_test

void ExImage::FromFile()
{
    //ExStart
    //ExFor:Shape.#ctor(DocumentBase,ShapeType)
    //ExFor:ShapeType
    //ExSummary:Shows how to insert a shape with an image from the local file system into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // The "Shape" class's public constructor will create a shape with "ShapeMarkupLanguage.Vml" markup type.
    // If you need to create a shape of a non-primitive type, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
    // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, or DiagonalCornersRounded,
    // please use DocumentBuilder.InsertShape.
    auto shape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Image);
    shape->get_ImageData()->SetImage(get_ImageDir() + u"Windows MetaFile.wmf");
    shape->set_Width(100);
    shape->set_Height(100);
    
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(shape);
    
    doc->Save(get_ArtifactsDir() + u"Image.FromFile.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Image.FromFile.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(1600, 1600, Aspose::Words::Drawing::ImageType::Wmf, shape);
    ASPOSE_ASSERT_EQ(100.0, shape->get_Height());
    ASPOSE_ASSERT_EQ(100.0, shape->get_Width());
}

namespace gtest_test
{

TEST_F(ExImage, FromFile)
{
    s_instance->FromFile();
}

} // namespace gtest_test

void ExImage::FromUrl()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(String)
    //ExSummary:Shows how to insert a shape with an image into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are two locations where the document builder's "InsertShape" method
    // can source the image that the shape will display.
    // 1 -  Pass a local file system filename of an image file:
    builder->Write(u"Image from local file: ");
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    builder->Writeln();
    
    // 2 -  Pass a URL which points to an image.
    builder->Write(u"Image from a URL: ");
    builder->InsertImage(get_ImageUrl());
    builder->Writeln();
    
    doc->Save(get_ArtifactsDir() + u"Image.FromUrl.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Image.FromUrl.docx");
    System::SharedPtr<Aspose::Words::NodeCollection> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);
    
    ASSERT_EQ(2, shapes->get_Count());
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, System::ExplicitCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(0)));
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(272, 92, Aspose::Words::Drawing::ImageType::Png, System::ExplicitCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(1)));
}

namespace gtest_test
{

TEST_F(ExImage, IgnoreOnJenkins_FromUrl)
{
    RecordProperty("category", "IgnoreOnJenkins");
    
    s_instance->FromUrl();
}

} // namespace gtest_test

void ExImage::FromStream()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(Stream)
    //ExSummary:Shows how to insert a shape with an image from a stream into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    {
        System::SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(get_ImageDir() + u"Logo.jpg");
        builder->Write(u"Image from stream: ");
        builder->InsertImage(stream);
    }
    
    doc->Save(get_ArtifactsDir() + u"Image.FromStream.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Image.FromStream.docx");
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->idx_get(0)));
}

namespace gtest_test
{

TEST_F(ExImage, FromStream)
{
    s_instance->FromStream();
}

} // namespace gtest_test

void ExImage::CreateFloatingPageCenter()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(String)
    //ExFor:Shape
    //ExFor:ShapeBase
    //ExFor:ShapeBase.WrapType
    //ExFor:ShapeBase.BehindText
    //ExFor:ShapeBase.RelativeHorizontalPosition
    //ExFor:ShapeBase.RelativeVerticalPosition
    //ExFor:ShapeBase.HorizontalAlignment
    //ExFor:ShapeBase.VerticalAlignment
    //ExFor:WrapType
    //ExFor:RelativeHorizontalPosition
    //ExFor:RelativeVerticalPosition
    //ExFor:HorizontalAlignment
    //ExFor:VerticalAlignment
    //ExSummary:Shows how to insert a floating image to the center of a page.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a floating image that will appear behind the overlapping text and align it to the page's center.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    shape->set_BehindText(true);
    shape->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::Page);
    shape->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::Page);
    shape->set_HorizontalAlignment(Aspose::Words::Drawing::HorizontalAlignment::Center);
    shape->set_VerticalAlignment(Aspose::Words::Drawing::VerticalAlignment::Center);
    
    doc->Save(get_ArtifactsDir() + u"Image.CreateFloatingPageCenter.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Image.CreateFloatingPageCenter.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, shape);
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::None, shape->get_WrapType());
    ASSERT_TRUE(shape->get_BehindText());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Page, shape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Page, shape->get_RelativeVerticalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::HorizontalAlignment::Center, shape->get_HorizontalAlignment());
    ASSERT_EQ(Aspose::Words::Drawing::VerticalAlignment::Center, shape->get_VerticalAlignment());
}

namespace gtest_test
{

TEST_F(ExImage, CreateFloatingPageCenter)
{
    s_instance->CreateFloatingPageCenter();
}

} // namespace gtest_test

void ExImage::CreateFloatingPositionSize()
{
    //ExStart
    //ExFor:ShapeBase.Left
    //ExFor:ShapeBase.Right
    //ExFor:ShapeBase.Top
    //ExFor:ShapeBase.Bottom
    //ExFor:ShapeBase.Width
    //ExFor:ShapeBase.Height
    //ExFor:DocumentBuilder.CurrentSection
    //ExFor:PageSetup.PageWidth
    //ExSummary:Shows how to insert a floating image, and specify its position and size.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    
    // Configure the shape's "RelativeHorizontalPosition" property to treat the value of the "Left" property
    // as the shape's horizontal distance, in points, from the left side of the page. 
    shape->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::Page);
    
    // Set the shape's horizontal distance from the left side of the page to 100.
    shape->set_Left(100);
    
    // Use the "RelativeVerticalPosition" property in a similar way to position the shape 80pt below the top of the page.
    shape->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::Page);
    shape->set_Top(80);
    
    // Set the shape's height, which will automatically scale the width to preserve dimensions.
    shape->set_Height(125);
    
    ASPOSE_ASSERT_EQ(125.0, shape->get_Width());
    
    // The "Bottom" and "Right" properties contain the bottom and right edges of the image.
    ASPOSE_ASSERT_EQ(shape->get_Top() + shape->get_Height(), shape->get_Bottom());
    ASPOSE_ASSERT_EQ(shape->get_Left() + shape->get_Width(), shape->get_Right());
    
    doc->Save(get_ArtifactsDir() + u"Image.CreateFloatingPositionSize.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Image.CreateFloatingPositionSize.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, shape);
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::None, shape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Page, shape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Page, shape->get_RelativeVerticalPosition());
    ASPOSE_ASSERT_EQ(100.0, shape->get_Left());
    ASPOSE_ASSERT_EQ(80.0, shape->get_Top());
    ASPOSE_ASSERT_EQ(125.0, shape->get_Height());
    ASPOSE_ASSERT_EQ(125.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(shape->get_Top() + shape->get_Height(), shape->get_Bottom());
    ASPOSE_ASSERT_EQ(shape->get_Left() + shape->get_Width(), shape->get_Right());
}

namespace gtest_test
{

TEST_F(ExImage, CreateFloatingPositionSize)
{
    s_instance->CreateFloatingPositionSize();
}

} // namespace gtest_test

void ExImage::InsertImageWithHyperlink()
{
    //ExStart
    //ExFor:ShapeBase.HRef
    //ExFor:ShapeBase.ScreenTip
    //ExFor:ShapeBase.Target
    //ExSummary:Shows how to insert a shape which contains an image, and is also a hyperlink.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    shape->set_HRef(u"https://forum.aspose.com/");
    shape->set_Target(u"New Window");
    shape->set_ScreenTip(u"Aspose.Words Support Forums");
    
    // Ctrl + left-clicking the shape in Microsoft Word will open a new web browser window
    // and take us to the hyperlink in the "HRef" property.
    doc->Save(get_ArtifactsDir() + u"Image.InsertImageWithHyperlink.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Image.InsertImageWithHyperlink.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(u"https://forum.aspose.com/", shape->get_HRef());
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, shape);
    ASSERT_EQ(u"New Window", shape->get_Target());
    ASSERT_EQ(u"Aspose.Words Support Forums", shape->get_ScreenTip());
}

namespace gtest_test
{

TEST_F(ExImage, InsertImageWithHyperlink)
{
    s_instance->InsertImageWithHyperlink();
}

} // namespace gtest_test

void ExImage::CreateLinkedImage()
{
    //ExStart
    //ExFor:Shape.ImageData
    //ExFor:ImageData
    //ExFor:ImageData.SourceFullName
    //ExFor:ImageData.SetImage(String)
    //ExFor:DocumentBuilder.InsertNode
    //ExSummary:Shows how to insert a linked image into a document. 
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::String imageFileName = get_ImageDir() + u"Windows MetaFile.wmf";
    
    // Below are two ways of applying an image to a shape so that it can display it.
    // 1 -  Set the shape to contain the image.
    auto shape = System::MakeObject<Aspose::Words::Drawing::Shape>(builder->get_Document(), Aspose::Words::Drawing::ShapeType::Image);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Inline);
    shape->get_ImageData()->SetImage(imageFileName);
    
    builder->InsertNode(shape);
    
    doc->Save(get_ArtifactsDir() + u"Image.CreateLinkedImage.Embedded.docx");
    
    // Every image that we store in shape will increase the size of our document.
    ASSERT_TRUE(70000 < System::MakeObject<System::IO::FileInfo>(get_ArtifactsDir() + u"Image.CreateLinkedImage.Embedded.docx")->get_Length());
    
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->RemoveAllChildren();
    
    // 2 -  Set the shape to link to an image file in the local file system.
    shape = System::MakeObject<Aspose::Words::Drawing::Shape>(builder->get_Document(), Aspose::Words::Drawing::ShapeType::Image);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Inline);
    shape->get_ImageData()->set_SourceFullName(imageFileName);
    
    builder->InsertNode(shape);
    doc->Save(get_ArtifactsDir() + u"Image.CreateLinkedImage.Linked.docx");
    
    // Linking to images will save space and result in a smaller document.
    // However, the document can only display the image correctly while
    // the image file is present at the location that the shape's "SourceFullName" property points to.
    ASSERT_TRUE(10000 > System::MakeObject<System::IO::FileInfo>(get_ArtifactsDir() + u"Image.CreateLinkedImage.Linked.docx")->get_Length());
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Image.CreateLinkedImage.Embedded.docx");
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(1600, 1600, Aspose::Words::Drawing::ImageType::Wmf, shape);
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, shape->get_WrapType());
    ASSERT_EQ(System::String::Empty, shape->get_ImageData()->get_SourceFullName().Replace(u"%20", u" "));
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Image.CreateLinkedImage.Linked.docx");
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(0, 0, Aspose::Words::Drawing::ImageType::Wmf, shape);
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, shape->get_WrapType());
    ASSERT_EQ(imageFileName, shape->get_ImageData()->get_SourceFullName().Replace(u"%20", u" "));
}

namespace gtest_test
{

TEST_F(ExImage, CreateLinkedImage)
{
    s_instance->CreateLinkedImage();
}

} // namespace gtest_test

void ExImage::DeleteAllImages()
{
    //ExStart
    //ExFor:Shape.HasImage
    //ExFor:Node.Remove
    //ExSummary:Shows how to delete all shapes with images from a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Images.docx");
    System::SharedPtr<Aspose::Words::NodeCollection> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);
    
    ASSERT_EQ(9, shapes->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Drawing::Shape>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Drawing::Shape> s)>>([](System::SharedPtr<Aspose::Words::Drawing::Shape> s) -> bool
    {
        return s->get_HasImage();
    }))));
    
    for (auto&& shape : System::IterateOver(shapes->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()))
    {
        if (shape->get_HasImage())
        {
            shape->Remove();
        }
    }
    
    ASSERT_EQ(0, shapes->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Drawing::Shape>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Drawing::Shape> s)>>([](System::SharedPtr<Aspose::Words::Drawing::Shape> s) -> bool
    {
        return s->get_HasImage();
    }))));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExImage, DeleteAllImages)
{
    s_instance->DeleteAllImages();
}

} // namespace gtest_test

void ExImage::DeleteAllImagesPreOrder()
{
    //ExStart
    //ExFor:Node.NextPreOrder(Node)
    //ExFor:Node.PreviousPreOrder(Node)
    //ExSummary:Shows how to traverse the document's node tree using the pre-order traversal algorithm, and delete any encountered shape with an image.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Images.docx");
    
    ASSERT_EQ(9, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Drawing::Shape>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Drawing::Shape> s)>>([](System::SharedPtr<Aspose::Words::Drawing::Shape> s) -> bool
    {
        return s->get_HasImage();
    }))));
    
    System::SharedPtr<Aspose::Words::Node> curNode = doc;
    while (curNode != nullptr)
    {
        System::SharedPtr<Aspose::Words::Node> nextNode = curNode->NextPreOrder(doc);
        
        if (curNode->PreviousPreOrder(doc) != nullptr && nextNode != nullptr)
        {
            ASPOSE_ASSERT_EQ(curNode, nextNode->PreviousPreOrder(doc));
        }
        
        if (curNode->get_NodeType() == Aspose::Words::NodeType::Shape && (System::ExplicitCast<Aspose::Words::Drawing::Shape>(curNode))->get_HasImage())
        {
            curNode->Remove();
        }
        
        curNode = nextNode;
    }
    
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Drawing::Shape>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Drawing::Shape> s)>>([](System::SharedPtr<Aspose::Words::Drawing::Shape> s) -> bool
    {
        return s->get_HasImage();
    }))));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExImage, DeleteAllImagesPreOrder)
{
    s_instance->DeleteAllImagesPreOrder();
}

} // namespace gtest_test

void ExImage::ScaleImage()
{
    //ExStart
    //ExFor:ImageData.ImageSize
    //ExFor:ImageSize
    //ExFor:ImageSize.WidthPoints
    //ExFor:ImageSize.HeightPoints
    //ExFor:ShapeBase.Width
    //ExFor:ShapeBase.Height
    //ExSummary:Shows how to resize a shape with an image.
    // When we insert an image using the "InsertImage" method, the builder scales the shape that displays the image so that,
    // when we view the document using 100% zoom in Microsoft Word, the shape displays the image in its actual size.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    // A 400x400 image will create an ImageData object with an image size of 300x300pt.
    System::SharedPtr<Aspose::Words::Drawing::ImageSize> imageSize = shape->get_ImageData()->get_ImageSize();
    
    ASPOSE_ASSERT_EQ(300.0, imageSize->get_WidthPoints());
    ASPOSE_ASSERT_EQ(300.0, imageSize->get_HeightPoints());
    
    // If a shape's dimensions match the image data's dimensions,
    // then the shape is displaying the image in its original size.
    ASPOSE_ASSERT_EQ(300.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(300.0, shape->get_Height());
    
    // Reduce the overall size of the shape by 50%. 
    System::WithLambda::setter_mul_wrap(GETTER_SETTER_LAMBDA_ARGS(shape, Width), 0.5);
    
    // Scaling factors apply to both the width and the height at the same time to preserve the shape's proportions. 
    ASPOSE_ASSERT_EQ(150.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(150.0, shape->get_Height());
    
    // When we resize the shape, the size of the image data remains the same.
    ASPOSE_ASSERT_EQ(300.0, imageSize->get_WidthPoints());
    ASPOSE_ASSERT_EQ(300.0, imageSize->get_HeightPoints());
    
    // We can reference the image data dimensions to apply a scaling based on the size of the image.
    shape->set_Width(imageSize->get_WidthPoints() * 1.1);
    
    ASPOSE_ASSERT_EQ(330.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(330.0, shape->get_Height());
    
    doc->Save(get_ArtifactsDir() + u"Image.ScaleImage.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Image.ScaleImage.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASPOSE_ASSERT_EQ(330.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(330.0, shape->get_Height());
    
    imageSize = shape->get_ImageData()->get_ImageSize();
    
    ASPOSE_ASSERT_EQ(300.0, imageSize->get_WidthPoints());
    ASPOSE_ASSERT_EQ(300.0, imageSize->get_HeightPoints());
}

namespace gtest_test
{

TEST_F(ExImage, ScaleImage)
{
    s_instance->ScaleImage();
}

} // namespace gtest_test

void ExImage::InsertWebpImage()
{
    //ExStart:InsertWebpImage
    //GistId:e386727403c2341ce4018bca370a5b41
    //ExFor:DocumentBuilder.InsertImage(String)
    //ExSummary:Shows how to insert WebP image.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertImage(get_ImageDir() + u"WebP image.webp");
    
    doc->Save(get_ArtifactsDir() + u"Image.InsertWebpImage.docx");
    //ExEnd:InsertWebpImage
}

namespace gtest_test
{

TEST_F(ExImage, InsertWebpImage)
{
    s_instance->InsertWebpImage();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
