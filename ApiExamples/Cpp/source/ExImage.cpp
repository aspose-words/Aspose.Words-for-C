// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExImage.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/file_stream.h>
#include <system/io/file.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/collections/list.h>
#include <system/collections/ienumerable.h>
#include <net/http_status_code.h>
#include <gtest/gtest.h>
#include <drawing/image.h>
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

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
namespace ApiExamples {

namespace gtest_test
{

class ExImage : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExImage> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExImage>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExImage> ExImage::s_instance;

} // namespace gtest_test

void ExImage::CreateImageDirectly()
{
    //ExStart
    //ExFor:Shape.#ctor(DocumentBase,ShapeType)
    //ExFor:ShapeType
    //ExSummary:Shows how to add a shape with an image to a document.
    auto doc = MakeObject<Document>();

    // Public constructor of "Shape" class creates shape with "ShapeMarkupLanguage.Vml" markup type
    // If you need to create non-primitive shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
    // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
    // please use DocumentBuilder.InsertShape
    auto shape = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Image);
    shape->get_ImageData()->SetImage(ImageDir + u"Windows MetaFile.wmf");
    shape->set_Width(100);
    shape->set_Height(100);

    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(shape);

    doc->Save(ArtifactsDir + u"Image.CreateImageDirectly.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Image.CreateImageDirectly.docx");
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyImageInShape(1600, 1600, Aspose::Words::Drawing::ImageType::Wmf, shape);
    ASPOSE_ASSERT_EQ(100.0, shape->get_Height());
    ASPOSE_ASSERT_EQ(100.0, shape->get_Width());
}

namespace gtest_test
{

TEST_F(ExImage, CreateImageDirectly)
{
    s_instance->CreateImageDirectly();
}

} // namespace gtest_test

void ExImage::CreateFromUrl()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(String)
    //ExSummary:Shows how to inserts an image from a URL.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Write(u"Image from local file: ");
    builder->InsertImage(ImageDir + u"Logo.jpg");
    builder->Writeln();

    builder->Write(u"Image from a URL: ");
    builder->InsertImage(AsposeLogoUrl);
    builder->Writeln();

    doc->Save(ArtifactsDir + u"Image.CreateFromUrl.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Image.CreateFromUrl.docx");
    SharedPtr<NodeCollection> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);

    ASSERT_EQ(2, shapes->get_Count());
    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, System::DynamicCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(0)));
    TestUtil::VerifyImageInShape(320, 320, Aspose::Words::Drawing::ImageType::Png, System::DynamicCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(1)));
}

namespace gtest_test
{

TEST_F(ExImage, CreateFromUrl)
{
    s_instance->CreateFromUrl();
}

} // namespace gtest_test

void ExImage::CreateFromStream()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(Stream)
    //ExSummary:Shows how to insert an image from a stream.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    {
        SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(ImageDir + u"Logo.jpg");
        builder->Write(u"Image from stream: ");
        builder->InsertImage(stream);
    }

    doc->Save(ArtifactsDir + u"Image.CreateFromStream.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Image.CreateFromStream.docx");

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->idx_get(0)));
}

namespace gtest_test
{

TEST_F(ExImage, CreateFromStream)
{
    s_instance->CreateFromStream();
}

} // namespace gtest_test

void ExImage::CreateFromImage()
{
    // This creates a builder and also an empty document inside the builder
    auto builder = MakeObject<DocumentBuilder>();

    // Insert a raster image
    {
        SharedPtr<System::Drawing::Image> rasterImage = System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg");
        builder->Write(u"Raster image: ");
        builder->InsertImage(rasterImage);
        builder->Writeln();
    }

    // Aspose.Words allows to insert a metafile too
    {
        SharedPtr<System::Drawing::Image> metafile = System::Drawing::Image::FromFile(ImageDir + u"Windows MetaFile.wmf");
        builder->Write(u"Metafile: ");
        builder->InsertImage(metafile);
        builder->Writeln();
    }

    builder->get_Document()->Save(ArtifactsDir + u"Image.CreateFromImage.docx");
}

namespace gtest_test
{

TEST_F(ExImage, SkipMono_CreateFromImage)
{
    RecordProperty("category", "SkipMono");

    s_instance->CreateFromImage();
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
    //ExSummary:Shows how to insert a floating image in the middle of a page.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // By default, the image is inline
    SharedPtr<Shape> shape = builder->InsertImage(ImageDir + u"Logo.jpg");

    // Make the image float, put it behind text and center on the page
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    shape->set_BehindText(true);
    shape->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::Page);
    shape->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::Page);
    shape->set_HorizontalAlignment(Aspose::Words::Drawing::HorizontalAlignment::Center);
    shape->set_VerticalAlignment(Aspose::Words::Drawing::VerticalAlignment::Center);

    doc->Save(ArtifactsDir + u"Image.CreateFloatingPageCenter.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Image.CreateFloatingPageCenter.docx");
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, shape);
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
    //ExSummary:Shows how to insert a floating image and specify its position and size.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // By default, the image is inline
    SharedPtr<Shape> shape = builder->InsertImage(ImageDir + u"Logo.jpg");

    // Make the image float, put it behind text and center on the page
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::None);

    // Make position relative to the page
    shape->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::Page);
    shape->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::Page);

    // Set the shape's coordinates, from the top left corner of the page
    shape->set_Left(100);
    shape->set_Top(80);

    // Set the shape's height
    shape->set_Height(125.0);

    // The width will be scaled to the height and the dimensions of the real image
    ASPOSE_ASSERT_EQ(125.0, shape->get_Width());

    // The Bottom and Right members contain the locations of the bottom and right edges of the image
    ASPOSE_ASSERT_EQ(shape->get_Top() + shape->get_Height(), shape->get_Bottom());
    ASPOSE_ASSERT_EQ(shape->get_Left() + shape->get_Width(), shape->get_Right());

    doc->Save(ArtifactsDir + u"Image.CreateFloatingPositionSize.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Image.CreateFloatingPositionSize.docx");
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, shape);
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
    //ExSummary:Shows how to insert an image with a hyperlink.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<Shape> shape = builder->InsertImage(ImageDir + u"Windows MetaFile.wmf");
    shape->set_HRef(u"https://forum.aspose.com/");
    shape->set_Target(u"New Window");
    shape->set_ScreenTip(u"Aspose.Words Support Forums");

    doc->Save(ArtifactsDir + u"Image.InsertImageWithHyperlink.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Image.InsertImageWithHyperlink.docx");
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, shape->get_HRef());
    TestUtil::VerifyImageInShape(1600, 1600, Aspose::Words::Drawing::ImageType::Wmf, shape);
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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    String imageFileName = ImageDir + u"Windows MetaFile.wmf";

    builder->Write(u"Image linked, not stored in the document: ");

    auto shape = MakeObject<Shape>(builder->get_Document(), Aspose::Words::Drawing::ShapeType::Image);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Inline);
    shape->get_ImageData()->set_SourceFullName(imageFileName);

    builder->InsertNode(shape);
    builder->Writeln();

    builder->Write(u"Image linked and stored in the document: ");

    shape = MakeObject<Shape>(builder->get_Document(), Aspose::Words::Drawing::ShapeType::Image);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Inline);
    shape->get_ImageData()->set_SourceFullName(imageFileName);
    shape->get_ImageData()->SetImage(imageFileName);

    builder->InsertNode(shape);
    builder->Writeln();

    builder->Write(u"Image stored in the document, but not linked: ");

    shape = MakeObject<Shape>(builder->get_Document(), Aspose::Words::Drawing::ShapeType::Image);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Inline);
    shape->get_ImageData()->SetImage(imageFileName);

    builder->InsertNode(shape);
    builder->Writeln();

    doc->Save(ArtifactsDir + u"Image.CreateLinkedImage.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Image.CreateLinkedImage.docx");

    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    TestUtil::VerifyImageInShape(0, 0, Aspose::Words::Drawing::ImageType::Wmf, shape);
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, shape->get_WrapType());
    ASSERT_EQ(imageFileName, shape->get_ImageData()->get_SourceFullName());

    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));

    TestUtil::VerifyImageInShape(1600, 1600, Aspose::Words::Drawing::ImageType::Wmf, shape);
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, shape->get_WrapType());
    ASSERT_EQ(imageFileName, shape->get_ImageData()->get_SourceFullName());

    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));

    TestUtil::VerifyImageInShape(1600, 1600, Aspose::Words::Drawing::ImageType::Wmf, shape);
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, shape->get_WrapType());
    ASSERT_EQ(String::Empty, shape->get_ImageData()->get_SourceFullName());
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
    //ExSummary:Shows how to delete all images from a document.
    auto doc = MakeObject<Document>(MyDir + u"Images.docx");
    ASSERT_EQ(10, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());

    // Here we get all shapes from the document node, but you can do this for any smaller
    // node too, for example delete shapes from a single section or a paragraph
    SharedPtr<NodeCollection> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);

    // We cannot delete shape nodes while we enumerate through the collection
    // One solution is to add nodes that we want to delete to a temporary array and delete afterwards
    SharedPtr<System::Collections::Generic::List<SharedPtr<Shape>>> shapesToDelete = MakeObject<System::Collections::Generic::List<SharedPtr<Shape>>>();

    // Several shape types can have an image including image shapes and OLE objects
    for (auto shape : System::IterateOver(shapes->LINQ_OfType<SharedPtr<Shape> >()))
    {
        if (shape->get_HasImage())
        {
            shapesToDelete->Add(shape);
        }
    }

    // Now we can delete shapes
    for (auto shape : shapesToDelete)
    {
        shape->Remove();
    }

    // The only remaining shape doesn't have an image
    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    ASSERT_FALSE((System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_HasImage());
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
    //ExSummary:Shows how to delete all images from a document using pre-order tree traversal.
    auto doc = MakeObject<Document>(MyDir + u"Images.docx");
    ASSERT_EQ(10, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());

    SharedPtr<Node> curNode = doc;
    while (curNode != nullptr)
    {
        SharedPtr<Node> nextNode = curNode->NextPreOrder(doc);

        if (curNode->PreviousPreOrder(doc) != nullptr && nextNode != nullptr)
        {
            ASPOSE_ASSERT_EQ(curNode, nextNode->PreviousPreOrder(doc));
        }

        // Several shape types can have an image including image shapes and OLE objects
        if (curNode->get_NodeType() == Aspose::Words::NodeType::Shape && (System::DynamicCast<Aspose::Words::Drawing::Shape>(curNode))->get_HasImage())
        {
            curNode->Remove();
        }

        curNode = nextNode;
    }

    // The only remaining shape doesn't have an image
    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    ASSERT_FALSE((System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_HasImage());
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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // By default, the image is inserted at 100% scale
    SharedPtr<Shape> shape = builder->InsertImage(ImageDir + u"Logo.jpg");

    // Reduce the overall size of the shape by 50%
    shape->set_Width(shape->get_Width() * 0.5);
    shape->set_Height(shape->get_Height() * 0.5);

    ASPOSE_ASSERT_EQ(75.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(75.0, shape->get_Height());

    // However, we can also go back to the original image size and scale from there, for example, to 110%
    SharedPtr<ImageSize> imageSize = shape->get_ImageData()->get_ImageSize();
    shape->set_Width(imageSize->get_WidthPoints() * 1.1);
    shape->set_Height(imageSize->get_HeightPoints() * 1.1);

    ASPOSE_ASSERT_EQ(330.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(330.0, shape->get_Height());

    doc->Save(ArtifactsDir + u"Image.ScaleImage.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Image.ScaleImage.docx");
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

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

} // namespace ApiExamples
