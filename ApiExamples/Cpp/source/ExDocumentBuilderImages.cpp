// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentBuilderImages.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/io/stream.h>
#include <system/io/memory_stream.h>
#include <system/io/file_stream.h>
#include <system/io/file.h>
#include <system/exceptions.h>
#include <system/details/dispose_guard.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <drawing/imaging/image_format.h>
#include <drawing/image.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageSize.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>

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

class ExDocumentBuilderImages : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExDocumentBuilderImages> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExDocumentBuilderImages>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExDocumentBuilderImages> ExDocumentBuilderImages::s_instance;

} // namespace gtest_test

void ExDocumentBuilderImages::InsertImageFromStream()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(Stream)
    //ExFor:DocumentBuilder.InsertImage(Stream, Double, Double)
    //ExFor:DocumentBuilder.InsertImage(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows different solutions of how to import an image into a document from a stream.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    {
        SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(ImageDir + u"Logo.jpg");
        builder->Writeln(u"Inserted image from stream: ");
        builder->InsertImage(stream);

        builder->Writeln(u"\nInserted image from stream with a custom size: ");
        builder->InsertImage(stream, ConvertUtil::PixelToPoint(250), ConvertUtil::PixelToPoint(144));

        builder->Writeln(u"\nInserted image from stream using relative positions: ");
        builder->InsertImage(stream, Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100.0, 200.0, 100.0, Aspose::Words::Drawing::WrapType::Square);
    }

    doc->Save(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromStream.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromStream.docx");

    auto imageShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());

    imageShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));

    ASPOSE_ASSERT_EQ(108.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(187.5, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());

    imageShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));

    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(200.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Top());

    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilderImages, InsertImageFromStream)
{
    s_instance->InsertImageFromStream();
}

} // namespace gtest_test

void ExDocumentBuilderImages::InsertImageFromString()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(String)
    //ExFor:DocumentBuilder.InsertImage(String, Double, Double)
    //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows different solutions of how to import an image into a document from a string.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Writeln(u"\nInserted image from string: ");
    builder->InsertImage(ImageDir + u"Logo.jpg");

    builder->Writeln(u"\nInserted image from string with a custom size: ");
    builder->InsertImage(ImageDir + u"Transparent background logo.png", ConvertUtil::PixelToPoint(250), ConvertUtil::PixelToPoint(144));

    builder->Writeln(u"\nInserted image from string using relative positions: ");
    builder->InsertImage(ImageDir + u"Windows Metafile.wmf", Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100.0, 200.0, 100.0, Aspose::Words::Drawing::WrapType::Square);

    doc->Save(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromString.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromString.docx");

    auto imageShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());

    imageShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));

    ASPOSE_ASSERT_EQ(108.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(187.5, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Png, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());

    imageShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));

    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(200.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Top());

    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());

    TestUtil::VerifyImageInShape(1600, 1600, Aspose::Words::Drawing::ImageType::Wmf, imageShape);
    ASPOSE_ASSERT_EQ(400.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(400.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilderImages, InsertImageFromString)
{
    s_instance->InsertImageFromString();
}

} // namespace gtest_test

void ExDocumentBuilderImages::InsertImageFromImageClass()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
    //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows different solutions of how to import an image into a document from Image class.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg");

    builder->Writeln(u"\nInserted image from Image class: ");
    builder->InsertImage(image);

    builder->Writeln(u"\nInserted image from Image class with a custom size: ");
    builder->InsertImage(image, ConvertUtil::PixelToPoint(250), ConvertUtil::PixelToPoint(144));

    builder->Writeln(u"\nInserted image from Image class using relative positions: ");
    builder->InsertImage(image, Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100.0, 200.0, 100.0, Aspose::Words::Drawing::WrapType::Square);

    doc->Save(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromImageClass.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromImageClass.docx");

    auto imageShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());

    imageShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));

    ASPOSE_ASSERT_EQ(108.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(187.5, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());

    imageShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));

    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(200.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Top());

    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilderImages, InsertImageFromImageClass)
{
    s_instance->InsertImageFromImageClass();
}

} // namespace gtest_test

void ExDocumentBuilderImages::InsertImageFromByteArray()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(Byte[])
    //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
    //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows different solutions of how to import an image into a document from a byte array.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg");

    {
        auto ms = MakeObject<System::IO::MemoryStream>();
        image->Save(ms, System::Drawing::Imaging::ImageFormat::get_Png());
        ArrayPtr<uint8_t> imageByteArray = ms->ToArray();

        builder->Writeln(u"\nInserted image from byte array: ");
        builder->InsertImage(imageByteArray);

        builder->Writeln(u"\nInserted image from byte array with a custom size: ");
        builder->InsertImage(imageByteArray, ConvertUtil::PixelToPoint(250), ConvertUtil::PixelToPoint(144));

        builder->Writeln(u"\nInserted image from byte array using relative positions: ");
        builder->InsertImage(imageByteArray, Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100.0, 200.0, 100.0, Aspose::Words::Drawing::WrapType::Square);
    }

    doc->Save(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromByteArray.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilderImages.InsertImageFromByteArray.docx");

    auto imageShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASSERT_NEAR(300.0, imageShape->get_Height(), 0.1);
    ASSERT_NEAR(300.0, imageShape->get_Width(), 0.1);
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Png, imageShape);
    ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints(), 0.1);
    ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints(), 0.1);

    imageShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));

    ASPOSE_ASSERT_EQ(108.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(187.5, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());

    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Png, imageShape);
    ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints(), 0.1);
    ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints(), 0.1);

    imageShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));

    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(200.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Top());

    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());

    TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Png, imageShape);
    ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints(), 0.1);
    ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints(), 0.1);
}

namespace gtest_test
{

TEST_F(ExDocumentBuilderImages, InsertImageFromByteArray)
{
    s_instance->InsertImageFromByteArray();
}

} // namespace gtest_test

} // namespace ApiExamples
