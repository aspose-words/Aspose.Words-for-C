#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/ImageSize.h>
#include <Aspose.Words.Cpp/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <drawing/image.h>
#include <net/http_status_code.h>
#include <system/collections/ienumerable.h>
#include <system/details/dispose_guard.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/io/file.h>
#include <system/io/file_info.h>
#include <system/io/file_stream.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/setters_wrap.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace ApiExamples {

/// <summary>
/// Mostly scenarios that deal with image shapes.
/// </summary>
class ExImage : public ApiExampleBase
{
public:
    void FromFile()
    {
        //ExStart
        //ExFor:Shape.#ctor(DocumentBase,ShapeType)
        //ExFor:ShapeType
        //ExSummary:Shows how to insert a shape with an image from the local file system into a document.
        auto doc = MakeObject<Document>();

        // The "Shape" class's public constructor will create a shape with "ShapeMarkupLanguage.Vml" markup type.
        // If you need to create a shape of a non-primitive type, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
        // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, or DiagonalCornersRounded,
        // please use DocumentBuilder.InsertShape.
        auto shape = MakeObject<Shape>(doc, ShapeType::Image);
        shape->get_ImageData()->SetImage(ImageDir + u"Windows MetaFile.wmf");
        shape->set_Width(100);
        shape->set_Height(100);

        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(shape);

        doc->Save(ArtifactsDir + u"Image.FromFile.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Image.FromFile.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(1600, 1600, ImageType::Wmf, shape);
        ASPOSE_ASSERT_EQ(100.0, shape->get_Height());
        ASPOSE_ASSERT_EQ(100.0, shape->get_Width());
    }

    void FromUrl()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExSummary:Shows how to insert a shape with an image into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are two locations where the document builder's "InsertShape" method
        // can source the image that the shape will display.
        // 1 -  Pass a local file system filename of an image file:
        builder->Write(u"Image from local file: ");
        builder->InsertImage(ImageDir + u"Logo.jpg");
        builder->Writeln();

        // 2 -  Pass a URL which points to an image.
        builder->Write(u"Image from a URL: ");
        builder->InsertImage(ImageUrl);
        builder->Writeln();

        doc->Save(ArtifactsDir + u"Image.FromUrl.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Image.FromUrl.docx");
        SharedPtr<NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);

        ASSERT_EQ(2, shapes->get_Count());
        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, System::DynamicCast<Shape>(shapes->idx_get(0)));
        TestUtil::VerifyImageInShape(5184, 3456, ImageType::Jpeg, System::DynamicCast<Shape>(shapes->idx_get(1)));
    }

    void FromStream()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Stream)
        //ExSummary:Shows how to insert a shape with an image from a stream into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        {
            SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(ImageDir + u"Logo.jpg");
            builder->Write(u"Image from stream: ");
            builder->InsertImage(stream);
        }

        doc->Save(ArtifactsDir + u"Image.FromStream.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Image.FromStream.docx");

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, System::DynamicCast<Shape>(doc->GetChildNodes(NodeType::Shape, true)->idx_get(0)));
    }

    void FromImage()
    {
        auto builder = MakeObject<DocumentBuilder>();

        {
            SharedPtr<System::Drawing::Image> rasterImage = System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg");
            builder->Write(u"Raster image: ");
            builder->InsertImage(rasterImage);
            builder->Writeln();
        }

        {
            SharedPtr<System::Drawing::Image> metafile = System::Drawing::Image::FromFile(ImageDir + u"Windows MetaFile.wmf");
            builder->Write(u"Metafile: ");
            builder->InsertImage(metafile);
            builder->Writeln();
        }

        builder->get_Document()->Save(ArtifactsDir + u"Image.FromImage.docx");
    }

    void CreateFloatingPageCenter()
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
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a floating image that will appear behind the overlapping text and align it to the page's center.
        SharedPtr<Shape> shape = builder->InsertImage(ImageDir + u"Logo.jpg");
        shape->set_WrapType(WrapType::None);
        shape->set_BehindText(true);
        shape->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);
        shape->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);
        shape->set_HorizontalAlignment(HorizontalAlignment::Center);
        shape->set_VerticalAlignment(VerticalAlignment::Center);

        doc->Save(ArtifactsDir + u"Image.CreateFloatingPageCenter.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Image.CreateFloatingPageCenter.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, shape);
        ASSERT_EQ(WrapType::None, shape->get_WrapType());
        ASSERT_TRUE(shape->get_BehindText());
        ASSERT_EQ(RelativeHorizontalPosition::Page, shape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Page, shape->get_RelativeVerticalPosition());
        ASSERT_EQ(HorizontalAlignment::Center, shape->get_HorizontalAlignment());
        ASSERT_EQ(VerticalAlignment::Center, shape->get_VerticalAlignment());
    }

    void CreateFloatingPositionSize()
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
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertImage(ImageDir + u"Logo.jpg");
        shape->set_WrapType(WrapType::None);

        // Configure the shape's "RelativeHorizontalPosition" property to treat the value of the "Left" property
        // as the shape's horizontal distance, in points, from the left side of the page.
        shape->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);

        // Set the shape's horizontal distance from the left side of the page to 100.
        shape->set_Left(100);

        // Use the "RelativeVerticalPosition" property in a similar way to position the shape 80pt below the top of the page.
        shape->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);
        shape->set_Top(80);

        // Set the shape's height, which will automatically scale the width to preserve dimensions.
        shape->set_Height(125);

        ASPOSE_ASSERT_EQ(125.0, shape->get_Width());

        // The "Bottom" and "Right" properties contain the bottom and right edges of the image.
        ASPOSE_ASSERT_EQ(shape->get_Top() + shape->get_Height(), shape->get_Bottom());
        ASPOSE_ASSERT_EQ(shape->get_Left() + shape->get_Width(), shape->get_Right());

        doc->Save(ArtifactsDir + u"Image.CreateFloatingPositionSize.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Image.CreateFloatingPositionSize.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, shape);
        ASSERT_EQ(WrapType::None, shape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Page, shape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Page, shape->get_RelativeVerticalPosition());
        ASPOSE_ASSERT_EQ(100.0, shape->get_Left());
        ASPOSE_ASSERT_EQ(80.0, shape->get_Top());
        ASPOSE_ASSERT_EQ(125.0, shape->get_Height());
        ASPOSE_ASSERT_EQ(125.0, shape->get_Width());
        ASPOSE_ASSERT_EQ(shape->get_Top() + shape->get_Height(), shape->get_Bottom());
        ASPOSE_ASSERT_EQ(shape->get_Left() + shape->get_Width(), shape->get_Right());
    }

    void InsertImageWithHyperlink()
    {
        //ExStart
        //ExFor:ShapeBase.HRef
        //ExFor:ShapeBase.ScreenTip
        //ExFor:ShapeBase.Target
        //ExSummary:Shows how to insert a shape which contains an image, and is also a hyperlink.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertImage(ImageDir + u"Logo.jpg");
        shape->set_HRef(u"https://forum.aspose.com/");
        shape->set_Target(u"New Window");
        shape->set_ScreenTip(u"Aspose.Words Support Forums");

        // Ctrl + left-clicking the shape in Microsoft Word will open a new web browser window
        // and take us to the hyperlink in the "HRef" property.
        doc->Save(ArtifactsDir + u"Image.InsertImageWithHyperlink.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Image.InsertImageWithHyperlink.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, shape->get_HRef());
        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, shape);
        ASSERT_EQ(u"New Window", shape->get_Target());
        ASSERT_EQ(u"Aspose.Words Support Forums", shape->get_ScreenTip());
    }

    void CreateLinkedImage()
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

        // Below are two ways of applying an image to a shape so that it can display it.
        // 1 -  Set the shape to contain the image.
        auto shape = MakeObject<Shape>(builder->get_Document(), ShapeType::Image);
        shape->set_WrapType(WrapType::Inline);
        shape->get_ImageData()->SetImage(imageFileName);

        builder->InsertNode(shape);

        doc->Save(ArtifactsDir + u"Image.CreateLinkedImage.Embedded.docx");

        // Every image that we store in shape will increase the size of our document.
        ASSERT_TRUE(70000 < MakeObject<System::IO::FileInfo>(ArtifactsDir + u"Image.CreateLinkedImage.Embedded.docx")->get_Length());

        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->RemoveAllChildren();

        // 2 -  Set the shape to link to an image file in the local file system.
        shape = MakeObject<Shape>(builder->get_Document(), ShapeType::Image);
        shape->set_WrapType(WrapType::Inline);
        shape->get_ImageData()->set_SourceFullName(imageFileName);

        builder->InsertNode(shape);
        doc->Save(ArtifactsDir + u"Image.CreateLinkedImage.Linked.docx");

        // Linking to images will save space and result in a smaller document.
        // However, the document can only display the image correctly while
        // the image file is present at the location that the shape's "SourceFullName" property points to.
        ASSERT_TRUE(10000 > MakeObject<System::IO::FileInfo>(ArtifactsDir + u"Image.CreateLinkedImage.Linked.docx")->get_Length());
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Image.CreateLinkedImage.Embedded.docx");

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(1600, 1600, ImageType::Wmf, shape);
        ASSERT_EQ(WrapType::Inline, shape->get_WrapType());
        ASSERT_EQ(String::Empty, shape->get_ImageData()->get_SourceFullName().Replace(u"%20", u" "));

        doc = MakeObject<Document>(ArtifactsDir + u"Image.CreateLinkedImage.Linked.docx");

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(0, 0, ImageType::Wmf, shape);
        ASSERT_EQ(WrapType::Inline, shape->get_WrapType());
        ASSERT_EQ(imageFileName, shape->get_ImageData()->get_SourceFullName().Replace(u"%20", u" "));
    }

    void DeleteAllImages()
    {
        //ExStart
        //ExFor:Shape.HasImage
        //ExFor:Node.Remove
        //ExSummary:Shows how to delete all shapes with images from a document.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");
        SharedPtr<NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);

        ASSERT_EQ(9, shapes->LINQ_OfType<SharedPtr<Shape>>()->LINQ_Count([](SharedPtr<Shape> s) { return s->get_HasImage(); }));

        for (const auto& shape : System::IterateOver(shapes->LINQ_OfType<SharedPtr<Shape>>()))
        {
            if (shape->get_HasImage())
            {
                shape->Remove();
            }
        }

        ASSERT_EQ(0, shapes->LINQ_OfType<SharedPtr<Shape>>()->LINQ_Count([](SharedPtr<Shape> s) { return s->get_HasImage(); }));
        //ExEnd
    }

    void DeleteAllImagesPreOrder()
    {
        //ExStart
        //ExFor:Node.NextPreOrder(Node)
        //ExFor:Node.PreviousPreOrder(Node)
        //ExSummary:Shows how to traverse the document's node tree using the pre-order traversal algorithm, and delete any encountered shape with an image.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        ASSERT_EQ(9,
                  doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_Count([](SharedPtr<Shape> s) { return s->get_HasImage(); }));

        SharedPtr<Node> curNode = doc;
        while (curNode != nullptr)
        {
            SharedPtr<Node> nextNode = curNode->NextPreOrder(doc);

            if (curNode->PreviousPreOrder(doc) != nullptr && nextNode != nullptr)
            {
                ASPOSE_ASSERT_EQ(curNode, nextNode->PreviousPreOrder(doc));
            }

            if (curNode->get_NodeType() == NodeType::Shape && (System::DynamicCast<Shape>(curNode))->get_HasImage())
            {
                curNode->Remove();
            }

            curNode = nextNode;
        }

        ASSERT_EQ(0,
                  doc->GetChildNodes(NodeType::Shape, true)->LINQ_OfType<SharedPtr<Shape>>()->LINQ_Count([](SharedPtr<Shape> s) { return s->get_HasImage(); }));
        //ExEnd
    }

    void ScaleImage()
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
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        SharedPtr<Shape> shape = builder->InsertImage(ImageDir + u"Logo.jpg");

        // A 400x400 image will create an ImageData object with an image size of 300x300pt.
        SharedPtr<ImageSize> imageSize = shape->get_ImageData()->get_ImageSize();

        ASPOSE_ASSERT_EQ(300.0, imageSize->get_WidthPoints());
        ASPOSE_ASSERT_EQ(300.0, imageSize->get_HeightPoints());

        // If a shape's dimensions match the image data's dimensions,
        // then the shape is displaying the image in its original size.
        ASPOSE_ASSERT_EQ(300.0, shape->get_Width());
        ASPOSE_ASSERT_EQ(300.0, shape->get_Height());

        // Reduce the overall size of the shape by 50%.
        System::setter_mul_wrap(shape.GetPointer(), &ShapeBase::get_Width, &ShapeBase::set_Width, 0.5);

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

        doc->Save(ArtifactsDir + u"Image.ScaleImage.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Image.ScaleImage.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASPOSE_ASSERT_EQ(330.0, shape->get_Width());
        ASPOSE_ASSERT_EQ(330.0, shape->get_Height());

        imageSize = shape->get_ImageData()->get_ImageSize();

        ASPOSE_ASSERT_EQ(300.0, imageSize->get_WidthPoints());
        ASPOSE_ASSERT_EQ(300.0, imageSize->get_HeightPoints());
    }
};

} // namespace ApiExamples
