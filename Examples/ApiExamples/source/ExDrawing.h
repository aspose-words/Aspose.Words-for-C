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
#include <Aspose.Words.Cpp/Border.h>
#include <Aspose.Words.Cpp/BorderCollection.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Drawing/ArrowLength.h>
#include <Aspose.Words.Cpp/Drawing/ArrowType.h>
#include <Aspose.Words.Cpp/Drawing/ArrowWidth.h>
#include <Aspose.Words.Cpp/Drawing/DashStyle.h>
#include <Aspose.Words.Cpp/Drawing/EndCap.h>
#include <Aspose.Words.Cpp/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Drawing/FillType.h>
#include <Aspose.Words.Cpp/Drawing/FlipOrientation.h>
#include <Aspose.Words.Cpp/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/ImageSize.h>
#include <Aspose.Words.Cpp/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Drawing/JoinStyle.h>
#include <Aspose.Words.Cpp/Drawing/LayoutFlow.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Drawing/Stroke.h>
#include <Aspose.Words.Cpp/Drawing/TextBox.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/VisitorAction.h>
#include <drawing/color.h>
#include <drawing/image.h>
#include <drawing/image_format_converter.h>
#include <drawing/imaging/image_format.h>
#include <drawing/rotate_flip_type.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ienumerator.h>
#include <system/collections/list.h>
#include <system/details/dispose_guard.h>
#include <system/enum.h>
#include <system/enum_helpers.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/file_access.h>
#include <system/io/file_info.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/string_builder.h>
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

namespace ApiExamples {

class ExDrawing : public ApiExampleBase
{
public:
    void VariousShapes()
    {
        //ExStart
        //ExFor:Drawing.ArrowLength
        //ExFor:Drawing.ArrowType
        //ExFor:Drawing.ArrowWidth
        //ExFor:Drawing.DashStyle
        //ExFor:Drawing.EndCap
        //ExFor:Drawing.Fill.ForeColor
        //ExFor:Drawing.Fill.ImageBytes
        //ExFor:Drawing.Fill.Visible
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

        // Below are four examples of shapes that we can insert into our documents.
        // 1 -  Dotted, horizontal, half-transparent red line
        // with an arrow on the left end and a diamond on the right end:
        auto arrow = MakeObject<Shape>(doc, ShapeType::Line);
        arrow->set_Width(200);
        arrow->get_Stroke()->set_Color(System::Drawing::Color::get_Red());
        arrow->get_Stroke()->set_StartArrowType(ArrowType::Arrow);
        arrow->get_Stroke()->set_StartArrowLength(ArrowLength::Long);
        arrow->get_Stroke()->set_StartArrowWidth(ArrowWidth::Wide);
        arrow->get_Stroke()->set_EndArrowType(ArrowType::Diamond);
        arrow->get_Stroke()->set_EndArrowLength(ArrowLength::Long);
        arrow->get_Stroke()->set_EndArrowWidth(ArrowWidth::Wide);
        arrow->get_Stroke()->set_DashStyle(DashStyle::Dash);
        arrow->get_Stroke()->set_Opacity(0.5);

        ASSERT_EQ(JoinStyle::Miter, arrow->get_Stroke()->get_JoinStyle());

        builder->InsertNode(arrow);

        // 2 -  Thick black diagonal line with rounded ends:
        auto line = MakeObject<Shape>(doc, ShapeType::Line);
        line->set_Top(40);
        line->set_Width(200);
        line->set_Height(20);
        line->set_StrokeWeight(5.0);
        line->get_Stroke()->set_EndCap(EndCap::Round);

        builder->InsertNode(line);

        // 3 -  Arrow with a green fill:
        auto filledInArrow = MakeObject<Shape>(doc, ShapeType::Arrow);
        filledInArrow->set_Width(200);
        filledInArrow->set_Height(40);
        filledInArrow->set_Top(100);
        filledInArrow->get_Fill()->set_ForeColor(System::Drawing::Color::get_Green());
        filledInArrow->get_Fill()->set_Visible(true);

        builder->InsertNode(filledInArrow);

        // 4 -  Arrow with a flipped orientation filled in with the Aspose logo:
        auto filledInArrowImg = MakeObject<Shape>(doc, ShapeType::Arrow);
        filledInArrowImg->set_Width(200);
        filledInArrowImg->set_Height(40);
        filledInArrowImg->set_Top(160);
        filledInArrowImg->set_FlipOrientation(FlipOrientation::Both);

        ArrayPtr<uint8_t> imageBytes = System::IO::File::ReadAllBytes(ImageDir + u"Logo.jpg");

        {
            auto stream = MakeObject<System::IO::MemoryStream>(imageBytes);
            SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromStream(stream);
            // When we flip the orientation of our arrow, we also flip the image that the arrow contains.
            // Flip the image the other way to cancel this out before getting the shape to display it.
            image->RotateFlip(System::Drawing::RotateFlipType::RotateNoneFlipXY);

            filledInArrowImg->get_ImageData()->SetImage(image);
            filledInArrowImg->get_Stroke()->set_JoinStyle(JoinStyle::Round);

            builder->InsertNode(filledInArrowImg);
        }

        doc->Save(ArtifactsDir + u"Drawing.VariousShapes.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Drawing.VariousShapes.docx");

        ASSERT_EQ(4, doc->GetChildNodes(NodeType::Shape, true)->get_Count());

        arrow = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_EQ(ShapeType::Line, arrow->get_ShapeType());
        ASPOSE_ASSERT_EQ(200.0, arrow->get_Width());
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), arrow->get_Stroke()->get_Color().ToArgb());
        ASSERT_EQ(ArrowType::Arrow, arrow->get_Stroke()->get_StartArrowType());
        ASSERT_EQ(ArrowLength::Long, arrow->get_Stroke()->get_StartArrowLength());
        ASSERT_EQ(ArrowWidth::Wide, arrow->get_Stroke()->get_StartArrowWidth());
        ASSERT_EQ(ArrowType::Diamond, arrow->get_Stroke()->get_EndArrowType());
        ASSERT_EQ(ArrowLength::Long, arrow->get_Stroke()->get_EndArrowLength());
        ASSERT_EQ(ArrowWidth::Wide, arrow->get_Stroke()->get_EndArrowWidth());
        ASSERT_EQ(DashStyle::Dash, arrow->get_Stroke()->get_DashStyle());
        ASPOSE_ASSERT_EQ(0.5, arrow->get_Stroke()->get_Opacity());

        line = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        ASSERT_EQ(ShapeType::Line, line->get_ShapeType());
        ASPOSE_ASSERT_EQ(40.0, line->get_Top());
        ASPOSE_ASSERT_EQ(200.0, line->get_Width());
        ASPOSE_ASSERT_EQ(20.0, line->get_Height());
        ASPOSE_ASSERT_EQ(5.0, line->get_StrokeWeight());
        ASSERT_EQ(EndCap::Round, line->get_Stroke()->get_EndCap());

        filledInArrow = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 2, true));

        ASSERT_EQ(ShapeType::Arrow, filledInArrow->get_ShapeType());
        ASPOSE_ASSERT_EQ(200.0, filledInArrow->get_Width());
        ASPOSE_ASSERT_EQ(40.0, filledInArrow->get_Height());
        ASPOSE_ASSERT_EQ(100.0, filledInArrow->get_Top());
        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), filledInArrow->get_Fill()->get_ForeColor().ToArgb());
        ASSERT_TRUE(filledInArrow->get_Fill()->get_Visible());

        filledInArrowImg = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 3, true));

        ASSERT_EQ(ShapeType::Arrow, filledInArrowImg->get_ShapeType());
        ASPOSE_ASSERT_EQ(200.0, filledInArrowImg->get_Width());
        ASPOSE_ASSERT_EQ(40.0, filledInArrowImg->get_Height());
        ASPOSE_ASSERT_EQ(160.0, filledInArrowImg->get_Top());
        ASSERT_EQ(FlipOrientation::Both, filledInArrowImg->get_FlipOrientation());
    }

    void TypeOfImage()
    {
        //ExStart
        //ExFor:Drawing.ImageType
        //ExSummary:Shows how to add an image to a shape and check its type.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        ArrayPtr<uint8_t> imageBytes = System::IO::File::ReadAllBytes(ImageDir + u"Logo.jpg");

        {
            auto stream = MakeObject<System::IO::MemoryStream>(imageBytes);
            SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromStream(stream);

            // The image in the URL is a .gif. Inserting it into a document converts it into a .png.
            SharedPtr<Shape> imgShape = builder->InsertImage(image);
            ASSERT_EQ(ImageType::Jpeg, imgShape->get_ImageData()->get_ImageType());
        }

        //ExEnd
    }

    void FillSolid()
    {
        //ExStart
        //ExFor:Fill.Color()
        //ExFor:Fill.Color(Color)
        //ExSummary:Shows how to convert any of the fills back to solid fill.
        auto doc = MakeObject<Document>(MyDir + u"Two color gradient.docx");

        // Get Fill object for Font of the first Run.
        SharedPtr<Fill> fill = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0)->get_Font()->get_Fill();

        // Check Fill properties of the Font.
        std::cout << String::Format(u"The type of the fill is: {0}", fill->get_FillType()) << std::endl;
        std::cout << "The foreground color of the fill is: " << fill->get_ForeColor() << std::endl;
        std::cout << "The fill is transparent at " << (fill->get_Transparency() * 100) << "%" << std::endl;

        // Change type of the fill to Solid with uniform green color.
        fill->Solid(System::Drawing::Color::get_Green());
        std::cout << "\nThe fill is changed:" << std::endl;
        std::cout << String::Format(u"The type of the fill is: {0}", fill->get_FillType()) << std::endl;
        std::cout << "The foreground color of the fill is: " << fill->get_ForeColor() << std::endl;
        std::cout << "The fill transparency is " << (fill->get_Transparency() * 100) << "%" << std::endl;

        doc->Save(ArtifactsDir + u"Drawing.FillSolid.docx");
        //ExEnd
    }

    void SaveAllImages()
    {
        //ExStart
        //ExFor:ImageData.HasImage
        //ExFor:ImageData.ToImage
        //ExFor:ImageData.Save(Stream)
        //ExSummary:Shows how to save all images from a document to the file system.
        auto imgSourceDoc = MakeObject<Document>(MyDir + u"Images.docx");

        // Shapes with the "HasImage" flag set store and display all the document's images.
        auto shapesWithImages =
            imgSourceDoc->GetChildNodes(NodeType::Shape, true)->LINQ_Cast<SharedPtr<Shape>>()->LINQ_Where([](SharedPtr<Shape> s) { return s->get_HasImage(); });

        // Go through each shape and save its image.
        auto formatConverter = MakeObject<System::Drawing::ImageFormatConverter>();

        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Shape>>> enumerator = shapesWithImages->GetEnumerator();
            int shapeIndex = 0;

            while (enumerator->MoveNext())
            {
                SharedPtr<ImageData> imageData = enumerator->get_Current()->get_ImageData();
                SharedPtr<System::Drawing::Imaging::ImageFormat> format = imageData->ToImage()->get_RawFormat();
                String fileExtension = formatConverter->ConvertToString(format);

                {
                    SharedPtr<System::IO::FileStream> fileStream =
                        System::IO::File::Create(ArtifactsDir + String::Format(u"Drawing.SaveAllImages.{0}.{1}", ++shapeIndex, fileExtension));
                    imageData->Save(fileStream);
                }
            }
        }
        //ExEnd

        ArrayPtr<String> imageFileNames =
            System::IO::Directory::GetFiles(ArtifactsDir)
                ->LINQ_Where(static_cast<System::Func<String, bool>>(
                    static_cast<std::function<bool(String s)>>([](String s) -> bool { return s.StartsWith(ArtifactsDir + u"Drawing.SaveAllImages."); })))
                ->LINQ_OrderBy(static_cast<System::Func<String, String>>(static_cast<std::function<String(String s)>>([](String s) -> String { return s; })))
                ->LINQ_ToArray();
        SharedPtr<System::Collections::Generic::List<SharedPtr<System::IO::FileInfo>>> fileInfos =
            imageFileNames
                ->LINQ_Select(
                    static_cast<System::Func<String, SharedPtr<System::IO::FileInfo>>>(static_cast<std::function<SharedPtr<System::IO::FileInfo>(String s)>>(
                        [](String s) -> SharedPtr<System::IO::FileInfo> { return MakeObject<System::IO::FileInfo>(s); })))
                ->LINQ_ToList();

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

    void ImportImage()
    {
        //ExStart
        //ExFor:ImageData.SetImage(Image)
        //ExFor:ImageData.SetImage(Stream)
        //ExSummary:Shows how to display images from the local file system in a document.
        auto doc = MakeObject<Document>();

        // To display an image in a document, we will need to create a shape
        // which will contain an image, and then append it to the document's body.
        SharedPtr<Shape> imgShape;

        // Below are two ways of getting an image from a file in the local file system.
        // 1 -  Create an image object from an image file:
        {
            SharedPtr<System::Drawing::Image> srcImage = System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg");
            imgShape = MakeObject<Shape>(doc, ShapeType::Image);
            doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(imgShape);
            imgShape->get_ImageData()->SetImage(srcImage);
        }

        // 2 -  Open an image file from the local file system using a stream:
        {
            SharedPtr<System::IO::Stream> stream =
                MakeObject<System::IO::FileStream>(ImageDir + u"Logo.jpg", System::IO::FileMode::Open, System::IO::FileAccess::Read);
            imgShape = MakeObject<Shape>(doc, ShapeType::Image);
            doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(imgShape);
            imgShape->get_ImageData()->SetImage(stream);
            imgShape->set_Left(150.0f);
        }

        doc->Save(ArtifactsDir + u"Drawing.ImportImage.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Drawing.ImportImage.docx");

        ASSERT_EQ(2, doc->GetChildNodes(NodeType::Shape, true)->get_Count());

        imgShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, imgShape);
        ASPOSE_ASSERT_EQ(0.0, imgShape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, imgShape->get_Top());
        ASPOSE_ASSERT_EQ(300.0, imgShape->get_Height());
        ASPOSE_ASSERT_EQ(300.0, imgShape->get_Width());
        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, imgShape);

        imgShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, imgShape);
        ASPOSE_ASSERT_EQ(150.0, imgShape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, imgShape->get_Top());
        ASPOSE_ASSERT_EQ(300.0, imgShape->get_Height());
        ASPOSE_ASSERT_EQ(300.0, imgShape->get_Width());
    }

    void StrokePattern()
    {
        //ExStart
        //ExFor:Stroke.Color2
        //ExFor:Stroke.ImageBytes
        //ExSummary:Shows how to process shape stroke features.
        auto doc = MakeObject<Document>(MyDir + u"Shape stroke pattern border.docx");
        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        SharedPtr<Stroke> stroke = shape->get_Stroke();

        // Strokes can have two colors, which are used to create a pattern defined by two-tone image data.
        // Strokes with a single color do not use the Color2 property.
        ASPOSE_ASSERT_EQ(System::Drawing::Color::FromArgb(255, 128, 0, 0), stroke->get_Color());
        ASPOSE_ASSERT_EQ(System::Drawing::Color::FromArgb(255, 255, 255, 0), stroke->get_Color2());

        ASSERT_FALSE(stroke->get_ImageBytes() == nullptr);
        System::IO::File::WriteAllBytes(ArtifactsDir + u"Drawing.StrokePattern.png", stroke->get_ImageBytes());
        //ExEnd

        TestUtil::VerifyImage(8, 8, ArtifactsDir + u"Drawing.StrokePattern.png");
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitShapeEnd(Shape)
    //ExFor:DocumentVisitor.VisitShapeStart(Shape)
    //ExFor:DocumentVisitor.VisitGroupShapeEnd(GroupShape)
    //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
    //ExFor:Drawing.GroupShape
    //ExFor:Drawing.GroupShape.#ctor(DocumentBase)
    //ExFor:Drawing.GroupShape.Accept(DocumentVisitor)
    //ExFor:ShapeBase.IsGroup
    //ExFor:ShapeBase.ShapeType
    //ExSummary:Shows how to create a group of shapes, and print its contents using a document visitor.
    void GroupOfShapes()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // If you need to create "NonPrimitive" shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
        // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
        // please use DocumentBuilder.InsertShape methods.
        auto balloon = MakeObject<Shape>(doc, ShapeType::Balloon);
        balloon->set_Width(200);
        balloon->set_Height(200);
        balloon->get_Stroke()->set_Color(System::Drawing::Color::get_Red());

        auto cube = MakeObject<Shape>(doc, ShapeType::Cube);
        cube->set_Width(100);
        cube->set_Height(100);
        cube->get_Stroke()->set_Color(System::Drawing::Color::get_Blue());

        auto group = MakeObject<GroupShape>(doc);
        group->AppendChild(balloon);
        group->AppendChild(cube);

        ASSERT_TRUE(group->get_IsGroup());

        builder->InsertNode(group);

        auto printer = MakeObject<ExDrawing::ShapeGroupPrinter>();
        group->Accept(printer);

        std::cout << printer->GetText() << std::endl;
        TestGroupShapes(doc);
        //ExSkip
    }

    /// <summary>
    /// Prints the contents of a visited shape group to the console.
    /// </summary>
    class ShapeGroupPrinter : public DocumentVisitor
    {
    public:
        ShapeGroupPrinter()
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
        }

        String GetText()
        {
            return mBuilder->ToString();
        }

        VisitorAction VisitGroupShapeStart(SharedPtr<GroupShape> groupShape) override
        {
            mBuilder->AppendLine(u"Shape group started:");
            return VisitorAction::Continue;
        }

        VisitorAction VisitGroupShapeEnd(SharedPtr<GroupShape> groupShape) override
        {
            mBuilder->AppendLine(u"End of shape group");
            return VisitorAction::Continue;
        }

        VisitorAction VisitShapeStart(SharedPtr<Shape> shape) override
        {
            mBuilder->AppendLine(String(u"\tShape - ") + System::ObjectExt::ToString(shape->get_ShapeType()) + u":");
            mBuilder->AppendLine(String(u"\t\tWidth: ") + shape->get_Width());
            mBuilder->AppendLine(String(u"\t\tHeight: ") + shape->get_Height());
            mBuilder->AppendLine(String(u"\t\tStroke color: ") + shape->get_Stroke()->get_Color());
            mBuilder->AppendLine(String(u"\t\tFill color: ") + shape->get_Fill()->get_ForeColor());
            return VisitorAction::Continue;
        }

        VisitorAction VisitShapeEnd(SharedPtr<Shape> shape) override
        {
            mBuilder->AppendLine(u"\tEnd of shape");
            return VisitorAction::Continue;
        }

    private:
        SharedPtr<System::Text::StringBuilder> mBuilder;
    };
    //ExEnd

    static void TestGroupShapes(SharedPtr<Document> doc)
    {
        doc = DocumentHelper::SaveOpen(doc);
        auto shapes = System::DynamicCast<GroupShape>(doc->GetChild(NodeType::GroupShape, 0, true));

        ASSERT_EQ(2, shapes->get_ChildNodes()->get_Count());

        auto shape = System::DynamicCast<Shape>(shapes->get_ChildNodes()->idx_get(0));

        ASSERT_EQ(ShapeType::Balloon, shape->get_ShapeType());
        ASPOSE_ASSERT_EQ(200.0, shape->get_Width());
        ASPOSE_ASSERT_EQ(200.0, shape->get_Height());
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), shape->get_StrokeColor().ToArgb());

        shape = System::DynamicCast<Shape>(shapes->get_ChildNodes()->idx_get(1));

        ASSERT_EQ(ShapeType::Cube, shape->get_ShapeType());
        ASPOSE_ASSERT_EQ(100.0, shape->get_Width());
        ASPOSE_ASSERT_EQ(100.0, shape->get_Height());
        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), shape->get_StrokeColor().ToArgb());
    }

    void TextBox_()
    {
        //ExStart
        //ExFor:Drawing.LayoutFlow
        //ExSummary:Shows how to add text to a text box, and change its orientation
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto textbox = MakeObject<Shape>(doc, ShapeType::TextBox);
        textbox->set_Width(100);
        textbox->set_Height(100);
        textbox->get_TextBox()->set_LayoutFlow(LayoutFlow::BottomToTop);

        textbox->AppendChild(MakeObject<Paragraph>(doc));
        builder->InsertNode(textbox);

        builder->MoveTo(textbox->get_FirstParagraph());
        builder->Write(u"This text is flipped 90 degrees to the left.");

        doc->Save(ArtifactsDir + u"Drawing.TextBox.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Drawing.TextBox.docx");
        textbox = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_EQ(ShapeType::TextBox, textbox->get_ShapeType());
        ASPOSE_ASSERT_EQ(100.0, textbox->get_Width());
        ASPOSE_ASSERT_EQ(100.0, textbox->get_Height());
        ASSERT_EQ(LayoutFlow::BottomToTop, textbox->get_TextBox()->get_LayoutFlow());
        ASSERT_EQ(u"This text is flipped 90 degrees to the left.", textbox->GetText().Trim());
    }

    void GetDataFromImage()
    {
        //ExStart
        //ExFor:ImageData.ImageBytes
        //ExFor:ImageData.ToByteArray
        //ExFor:ImageData.ToStream
        //ExSummary:Shows how to create an image file from a shape's raw image data.
        auto imgSourceDoc = MakeObject<Document>(MyDir + u"Images.docx");
        ASSERT_EQ(10, imgSourceDoc->GetChildNodes(NodeType::Shape, true)->get_Count());
        //ExSkip

        auto imgShape = System::DynamicCast<Shape>(imgSourceDoc->GetChild(NodeType::Shape, 0, true));

        ASSERT_TRUE(imgShape->get_HasImage());

        // ToByteArray() returns the array stored in the ImageBytes property.
        ASPOSE_ASSERT_EQ(imgShape->get_ImageData()->get_ImageBytes(), imgShape->get_ImageData()->ToByteArray());

        // Save the shape's image data to an image file in the local file system.
        {
            SharedPtr<System::IO::Stream> imgStream = imgShape->get_ImageData()->ToStream();
            {
                auto outStream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"Drawing.GetDataFromImage.png", System::IO::FileMode::Create,
                                                                    System::IO::FileAccess::ReadWrite);
                imgStream->CopyTo(outStream);
            }
        }
        //ExEnd

        TestUtil::VerifyImage(2467, 1500, ArtifactsDir + u"Drawing.GetDataFromImage.png");
    }

    void ImageData_()
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
        auto imgSourceDoc = MakeObject<Document>(MyDir + u"Images.docx");
        auto sourceShape = System::DynamicCast<Shape>(imgSourceDoc->GetChildNodes(NodeType::Shape, true)->idx_get(0));

        auto dstDoc = MakeObject<Document>();

        // Import a shape from the source document and append it to the first paragraph.
        auto importedShape = System::DynamicCast<Shape>(dstDoc->ImportNode(sourceShape, true));
        dstDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(importedShape);

        // The imported shape contains an image. We can access the image's properties and raw data via the ImageData object.
        SharedPtr<ImageData> imageData = importedShape->get_ImageData();
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
        importedShape = System::DynamicCast<Shape>(dstDoc->ImportNode(sourceShape, true));
        dstDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(importedShape);

        importedShape->get_ImageData()->set_GrayScale(true);

        // Import the source shape again to create a third image and set it to BiLevel.
        // BiLevel sets every pixel to either black or white, whichever is closer to the original color.
        importedShape = System::DynamicCast<Shape>(dstDoc->ImportNode(sourceShape, true));
        dstDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(importedShape);

        importedShape->get_ImageData()->set_BiLevel(true);

        // Cropping is determined on a 0-1 scale. Cropping a side by 0.3
        // will crop 30% of the image out at the cropped side.
        importedShape->get_ImageData()->set_CropBottom(0.3);
        importedShape->get_ImageData()->set_CropLeft(0.3);
        importedShape->get_ImageData()->set_CropTop(0.3);
        importedShape->get_ImageData()->set_CropRight(0.3);

        dstDoc->Save(ArtifactsDir + u"Drawing.ImageData.docx");
        //ExEnd

        imgSourceDoc = MakeObject<Document>(ArtifactsDir + u"Drawing.ImageData.docx");
        sourceShape = System::DynamicCast<Shape>(imgSourceDoc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(2467, 1500, ImageType::Jpeg, sourceShape);
        ASSERT_EQ(u"Imported Image", sourceShape->get_ImageData()->get_Title());
        ASSERT_NEAR(0.8, sourceShape->get_ImageData()->get_Brightness(), 0.1);
        ASSERT_NEAR(1.0, sourceShape->get_ImageData()->get_Contrast(), 0.1);
        ASSERT_EQ(System::Drawing::Color::get_White().ToArgb(), sourceShape->get_ImageData()->get_ChromaKey().ToArgb());

        sourceShape = System::DynamicCast<Shape>(imgSourceDoc->GetChild(NodeType::Shape, 1, true));

        TestUtil::VerifyImageInShape(2467, 1500, ImageType::Jpeg, sourceShape);
        ASSERT_TRUE(sourceShape->get_ImageData()->get_GrayScale());

        sourceShape = System::DynamicCast<Shape>(imgSourceDoc->GetChild(NodeType::Shape, 2, true));

        TestUtil::VerifyImageInShape(2467, 1500, ImageType::Jpeg, sourceShape);
        ASSERT_TRUE(sourceShape->get_ImageData()->get_BiLevel());
        ASSERT_NEAR(0.3, sourceShape->get_ImageData()->get_CropBottom(), 0.1);
        ASSERT_NEAR(0.3, sourceShape->get_ImageData()->get_CropLeft(), 0.1);
        ASSERT_NEAR(0.3, sourceShape->get_ImageData()->get_CropTop(), 0.1);
        ASSERT_NEAR(0.3, sourceShape->get_ImageData()->get_CropRight(), 0.1);
    }

    void ImageSize_()
    {
        //ExStart
        //ExFor:ImageSize.HeightPixels
        //ExFor:ImageSize.HorizontalResolution
        //ExFor:ImageSize.VerticalResolution
        //ExFor:ImageSize.WidthPixels
        //ExSummary:Shows how to read the properties of an image in a shape.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a shape into the document which contains an image taken from our local file system.
        SharedPtr<Shape> shape = builder->InsertImage(ImageDir + u"Logo.jpg");

        // If the shape contains an image, its ImageData property will be valid,
        // and it will contain an ImageSize object.
        SharedPtr<ImageSize> imageSize = shape->get_ImageData()->get_ImageSize();

        // The ImageSize object contains read-only information about the image within the shape.
        ASSERT_EQ(400, imageSize->get_HeightPixels());
        ASSERT_EQ(400, imageSize->get_WidthPixels());

        const double delta = 0.05;
        ASSERT_NEAR(95.98, imageSize->get_HorizontalResolution(), delta);
        ASSERT_NEAR(95.98, imageSize->get_VerticalResolution(), delta);

        // We can base the size of the shape on the size of its image to avoid stretching the image.
        shape->set_Width(imageSize->get_WidthPoints() * 2);
        shape->set_Height(imageSize->get_HeightPoints() * 2);

        doc->Save(ArtifactsDir + u"Drawing.ImageSize.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Drawing.ImageSize.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, shape);
        ASPOSE_ASSERT_EQ(600.0, shape->get_Width());
        ASPOSE_ASSERT_EQ(600.0, shape->get_Height());

        imageSize = shape->get_ImageData()->get_ImageSize();

        ASSERT_EQ(400, imageSize->get_HeightPixels());
        ASSERT_EQ(400, imageSize->get_WidthPixels());
        ASSERT_NEAR(95.98, imageSize->get_HorizontalResolution(), delta);
        ASSERT_NEAR(95.98, imageSize->get_VerticalResolution(), delta);
    }
};

} // namespace ApiExamples
