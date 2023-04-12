#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/ConvertUtil.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/ImageSize.h>
#include <Aspose.Words.Cpp/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Settings/MsWordVersion.h>
#include <drawing/image.h>
#include <drawing/imaging/image_format.h>
#include <system/array.h>
#include <system/details/dispose_guard.h>
#include <system/exceptions.h>
#include <system/io/file.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/io/stream.h>
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
using namespace Aspose::Words::Settings;

namespace ApiExamples {

class ExDocumentBuilderImages : public ApiExampleBase
{
public:
    void InsertImageFromStream()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Stream)
        //ExFor:DocumentBuilder.InsertImage(Stream, Double, Double)
        //ExFor:DocumentBuilder.InsertImage(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image from a stream into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        {
            SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(ImageDir + u"Logo.jpg");
            // Below are three ways of inserting an image from a stream.
            // 1 -  Inline shape with a default size based on the image's original dimensions:
            builder->InsertImage(stream);

            builder->InsertBreak(BreakType::PageBreak);

            // 2 -  Inline shape with custom dimensions:
            builder->InsertImage(stream, ConvertUtil::PixelToPoint(250), ConvertUtil::PixelToPoint(144));

            builder->InsertBreak(BreakType::PageBreak);

            // 3 -  Floating shape with custom dimensions:
            builder->InsertImage(stream, RelativeHorizontalPosition::Margin, 100.0, RelativeVerticalPosition::Margin, 100.0, 200.0, 100.0, WrapType::Square);
        }

        doc->Save(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromStream.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromStream.docx");

        auto imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

        ASSERT_EQ(WrapType::Inline, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, imageShape);
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());

        imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        ASPOSE_ASSERT_EQ(108.0, imageShape->get_Height());
        ASPOSE_ASSERT_EQ(187.5, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

        ASSERT_EQ(WrapType::Inline, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, imageShape);
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());

        imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 2, true));

        ASPOSE_ASSERT_EQ(100.0, imageShape->get_Height());
        ASPOSE_ASSERT_EQ(200.0, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(100.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(100.0, imageShape->get_Top());

        ASSERT_EQ(WrapType::Square, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, imageShape);
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
    }

    void InsertImageFromFilename()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExFor:DocumentBuilder.InsertImage(String, Double, Double)
        //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image from the local file system into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are three ways of inserting an image from a local system filename.
        // 1 -  Inline shape with a default size based on the image's original dimensions:
        builder->InsertImage(ImageDir + u"Logo.jpg");

        builder->InsertBreak(BreakType::PageBreak);

        // 2 -  Inline shape with custom dimensions:
        builder->InsertImage(ImageDir + u"Transparent background logo.png", ConvertUtil::PixelToPoint(250), ConvertUtil::PixelToPoint(144));

        builder->InsertBreak(BreakType::PageBreak);

        // 3 -  Floating shape with custom dimensions:
        builder->InsertImage(ImageDir + u"Windows MetaFile.wmf", RelativeHorizontalPosition::Margin, 100.0, RelativeVerticalPosition::Margin, 100.0, 200.0,
                             100.0, WrapType::Square);

        doc->Save(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromFilename.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromFilename.docx");

        auto imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

        ASSERT_EQ(WrapType::Inline, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, imageShape);
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());

        imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        ASPOSE_ASSERT_EQ(108.0, imageShape->get_Height());
        ASPOSE_ASSERT_EQ(187.5, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

        ASSERT_EQ(WrapType::Inline, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

        TestUtil::VerifyImageInShape(400, 400, ImageType::Png, imageShape);
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());

        imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 2, true));

        ASPOSE_ASSERT_EQ(100.0, imageShape->get_Height());
        ASPOSE_ASSERT_EQ(200.0, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(100.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(100.0, imageShape->get_Top());

        ASSERT_EQ(WrapType::Square, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());

        TestUtil::VerifyImageInShape(1600, 1600, ImageType::Wmf, imageShape);
        ASPOSE_ASSERT_EQ(400.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
        ASPOSE_ASSERT_EQ(400.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
    }

    void InsertSvgImage()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExSummary:Shows how to determine which image will be inserted.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertImage(ImageDir + u"Scalable Vector Graphics.svg");

        // Aspose.Words insert SVG image to the document as PNG with svgBlip extension
        // that contains the original vector SVG image representation.
        doc->Save(ArtifactsDir + u"DocumentBuilderImages.InsertSvgImage.SvgWithSvgBlip.docx");

        // Aspose.Words insert SVG image to the document as PNG, just like Microsoft Word does for old format.
        doc->Save(ArtifactsDir + u"DocumentBuilderImages.InsertSvgImage.Svg.doc");

        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2003);

        // Aspose.Words insert SVG image to the document as EMF metafile to keep the image in vector representation.
        doc->Save(ArtifactsDir + u"DocumentBuilderImages.InsertSvgImage.Emf.docx");
        //ExEnd
    }

    void InsertImageFromImageObject()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Image)
        //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
        //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image from an object into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg");

        // Below are three ways of inserting an image from an Image object instance.
        // 1 -  Inline shape with a default size based on the image's original dimensions:
        builder->InsertImage(image);

        builder->InsertBreak(BreakType::PageBreak);

        // 2 -  Inline shape with custom dimensions:
        builder->InsertImage(image, ConvertUtil::PixelToPoint(250), ConvertUtil::PixelToPoint(144));

        builder->InsertBreak(BreakType::PageBreak);

        // 3 -  Floating shape with custom dimensions:
        builder->InsertImage(image, RelativeHorizontalPosition::Margin, 100.0, RelativeVerticalPosition::Margin, 100.0, 200.0, 100.0, WrapType::Square);

        doc->Save(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromImageObject.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromImageObject.docx");

        auto imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

        ASSERT_EQ(WrapType::Inline, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, imageShape);
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());

        imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        ASPOSE_ASSERT_EQ(108.0, imageShape->get_Height());
        ASPOSE_ASSERT_EQ(187.5, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

        ASSERT_EQ(WrapType::Inline, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, imageShape);
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());

        imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 2, true));

        ASPOSE_ASSERT_EQ(100.0, imageShape->get_Height());
        ASPOSE_ASSERT_EQ(200.0, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(100.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(100.0, imageShape->get_Top());

        ASSERT_EQ(WrapType::Square, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, imageShape);
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
        ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
    }

    void InsertImageFromByteArray()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Byte[])
        //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
        //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image from a byte array into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg");

        {
            auto ms = MakeObject<System::IO::MemoryStream>();
            image->Save(ms, System::Drawing::Imaging::ImageFormat::get_Png());
            ArrayPtr<uint8_t> imageByteArray = ms->ToArray();

            // Below are three ways of inserting an image from a byte array.
            // 1 -  Inline shape with a default size based on the image's original dimensions:
            builder->InsertImage(imageByteArray);

            builder->InsertBreak(BreakType::PageBreak);

            // 2 -  Inline shape with custom dimensions:
            builder->InsertImage(imageByteArray, ConvertUtil::PixelToPoint(250), ConvertUtil::PixelToPoint(144));

            builder->InsertBreak(BreakType::PageBreak);

            // 3 -  Floating shape with custom dimensions:
            builder->InsertImage(imageByteArray, RelativeHorizontalPosition::Margin, 100.0, RelativeVerticalPosition::Margin, 100.0, 200.0, 100.0,
                                 WrapType::Square);
        }

        doc->Save(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromByteArray.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromByteArray.docx");

        auto imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_NEAR(300.0, imageShape->get_Height(), 0.1);
        ASSERT_NEAR(300.0, imageShape->get_Width(), 0.1);
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

        ASSERT_EQ(WrapType::Inline, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

        TestUtil::VerifyImageInShape(400, 400, ImageType::Png, imageShape);
        ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints(), 0.1);
        ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints(), 0.1);

        imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        ASPOSE_ASSERT_EQ(108.0, imageShape->get_Height());
        ASPOSE_ASSERT_EQ(187.5, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

        ASSERT_EQ(WrapType::Inline, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

        TestUtil::VerifyImageInShape(400, 400, ImageType::Png, imageShape);
        ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints(), 0.1);
        ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints(), 0.1);

        imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 2, true));

        ASPOSE_ASSERT_EQ(100.0, imageShape->get_Height());
        ASPOSE_ASSERT_EQ(200.0, imageShape->get_Width());
        ASPOSE_ASSERT_EQ(100.0, imageShape->get_Left());
        ASPOSE_ASSERT_EQ(100.0, imageShape->get_Top());

        ASSERT_EQ(WrapType::Square, imageShape->get_WrapType());
        ASSERT_EQ(RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
        ASSERT_EQ(RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());

        TestUtil::VerifyImageInShape(400, 400, ImageType::Png, imageShape);
        ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints(), 0.1);
        ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints(), 0.1);
    }

    void InsertGif()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExSummary:Shows how to insert gif image to the document.
        auto builder = MakeObject<DocumentBuilder>();

        // We can insert gif image using path or bytes array.
        // It works only if DocumentBuilder optimized to Word version 2010 or higher.
        // Note, that access to the image bytes causes conversion Gif to Png.
        SharedPtr<Shape> gifImage = builder->InsertImage(ImageDir + u"Graphics Interchange Format.gif");

        gifImage = builder->InsertImage(System::IO::File::ReadAllBytes(ImageDir + u"Graphics Interchange Format.gif"));

        builder->get_Document()->Save(ArtifactsDir + u"InsertGif.docx");
        //ExEnd
    }
};

} // namespace ApiExamples
