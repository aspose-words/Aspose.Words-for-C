// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentBuilderImages.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/io/stream.h>
#include <system/io/file_stream.h>
#include <system/io/file.h>
#include <system/details/dispose_guard.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>
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
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

#include "TestUtil.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Settings;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1913370696u, ::Aspose::Words::ApiExamples::ExDocumentBuilderImages, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExDocumentBuilderImages : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExDocumentBuilderImages> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExDocumentBuilderImages>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExDocumentBuilderImages> ExDocumentBuilderImages::s_instance;

} // namespace gtest_test

void ExDocumentBuilderImages::InsertImageFromStream()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(Stream)
    //ExFor:DocumentBuilder.InsertImage(Stream, Double, Double)
    //ExFor:DocumentBuilder.InsertImage(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an image from a stream into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    {
        System::SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(get_ImageDir() + u"Logo.jpg");
        // Below are three ways of inserting an image from a stream.
        // 1 -  Inline shape with a default size based on the image's original dimensions:
        builder->InsertImage(stream);
        
        builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
        
        // 2 -  Inline shape with custom dimensions:
        builder->InsertImage(stream, Aspose::Words::ConvertUtil::PixelToPoint(250), Aspose::Words::ConvertUtil::PixelToPoint(144));
        
        builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
        
        // 3 -  Floating shape with custom dimensions:
        builder->InsertImage(stream, Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100.0, 200.0, 100.0, Aspose::Words::Drawing::WrapType::Square);
    }
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilderImages.InsertImageFromStream.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilderImages.InsertImageFromStream.docx");
    
    auto imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
    
    imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    ASPOSE_ASSERT_EQ(108.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(187.5, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
    
    imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));
    
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(200.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Top());
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
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

void ExDocumentBuilderImages::InsertImageFromFilename()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(String)
    //ExFor:DocumentBuilder.InsertImage(String, Double, Double)
    //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an image from the local file system into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are three ways of inserting an image from a local system filename.
    // 1 -  Inline shape with a default size based on the image's original dimensions:
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    // 2 -  Inline shape with custom dimensions:
    builder->InsertImage(get_ImageDir() + u"Transparent background logo.png", Aspose::Words::ConvertUtil::PixelToPoint(250), Aspose::Words::ConvertUtil::PixelToPoint(144));
    
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    // 3 -  Floating shape with custom dimensions:
    builder->InsertImage(get_ImageDir() + u"Windows MetaFile.wmf", Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100.0, 200.0, 100.0, Aspose::Words::Drawing::WrapType::Square);
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilderImages.InsertImageFromFilename.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilderImages.InsertImageFromFilename.docx");
    
    auto imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
    
    imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    ASPOSE_ASSERT_EQ(108.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(187.5, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Png, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
    
    imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));
    
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(200.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Top());
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(1600, 1600, Aspose::Words::Drawing::ImageType::Wmf, imageShape);
    ASPOSE_ASSERT_EQ(400.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(400.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilderImages, InsertImageFromFilename)
{
    s_instance->InsertImageFromFilename();
}

} // namespace gtest_test

void ExDocumentBuilderImages::InsertSvgImage()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(String)
    //ExSummary:Shows how to determine which image will be inserted.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertImage(get_ImageDir() + u"Scalable Vector Graphics.svg");
    
    // Aspose.Words insert SVG image to the document as PNG with svgBlip extension
    // that contains the original vector SVG image representation.
    doc->Save(get_ArtifactsDir() + u"DocumentBuilderImages.InsertSvgImage.SvgWithSvgBlip.docx");
    
    // Aspose.Words insert SVG image to the document as PNG, just like Microsoft Word does for old format.
    doc->Save(get_ArtifactsDir() + u"DocumentBuilderImages.InsertSvgImage.Svg.doc");
    
    doc->get_CompatibilityOptions()->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2003);
    
    // Aspose.Words insert SVG image to the document as EMF metafile to keep the image in vector representation.
    doc->Save(get_ArtifactsDir() + u"DocumentBuilderImages.InsertSvgImage.Emf.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilderImages, InsertSvgImage)
{
    s_instance->InsertSvgImage();
}

} // namespace gtest_test

void ExDocumentBuilderImages::InsertImageFromImageObject()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(Image)
    //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
    //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an image from an object into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::String imageFile = get_ImageDir() + u"Logo.jpg";
    
    // Below are three ways of inserting an image from an Image object instance.
    // 1 -  Inline shape with a default size based on the image's original dimensions:
    builder->InsertImage(imageFile);
    
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    // 2 -  Inline shape with custom dimensions:
    builder->InsertImage(imageFile, Aspose::Words::ConvertUtil::PixelToPoint(250), Aspose::Words::ConvertUtil::PixelToPoint(144));
    
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    // 3 -  Floating shape with custom dimensions:
    builder->InsertImage(imageFile, Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100.0, 200.0, 100.0, Aspose::Words::Drawing::WrapType::Square);
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilderImages.InsertImageFromImageObject.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilderImages.InsertImageFromImageObject.docx");
    
    auto imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
    
    imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    ASPOSE_ASSERT_EQ(108.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(187.5, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
    
    imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));
    
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(200.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Top());
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints());
    ASPOSE_ASSERT_EQ(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints());
}

namespace gtest_test
{

TEST_F(ExDocumentBuilderImages, InsertImageFromImageObject)
{
    s_instance->InsertImageFromImageObject();
}

} // namespace gtest_test

void ExDocumentBuilderImages::InsertImageFromByteArray()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(Byte[])
    //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
    //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an image from a byte array into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::ArrayPtr<uint8_t> imageByteArray = Aspose::Words::ApiExamples::TestUtil::ImageToByteArray(get_ImageDir() + u"Logo.jpg");
    
    // Below are three ways of inserting an image from a byte array.
    // 1 -  Inline shape with a default size based on the image's original dimensions:
    builder->InsertImage(imageByteArray);
    
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    // 2 -  Inline shape with custom dimensions:
    builder->InsertImage(imageByteArray, Aspose::Words::ConvertUtil::PixelToPoint(250), Aspose::Words::ConvertUtil::PixelToPoint(144));
    
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    // 3 -  Floating shape with custom dimensions:
    builder->InsertImage(imageByteArray, Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100.0, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100.0, 200.0, 100.0, Aspose::Words::Drawing::WrapType::Square);
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilderImages.InsertImageFromByteArray.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilderImages.InsertImageFromByteArray.docx");
    
    auto imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_NEAR(300.0, imageShape->get_Height(), 0.1);
    ASSERT_NEAR(300.0, imageShape->get_Width(), 0.1);
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints(), 0.1);
    ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints(), 0.1);
    
    imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    ASPOSE_ASSERT_EQ(108.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(187.5, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(0.0, imageShape->get_Top());
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Inline, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Column, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, imageShape->get_RelativeVerticalPosition());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_HeightPoints(), 0.1);
    ASSERT_NEAR(300.0, imageShape->get_ImageData()->get_ImageSize()->get_WidthPoints(), 0.1);
    
    imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 2, true));
    
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Height());
    ASPOSE_ASSERT_EQ(200.0, imageShape->get_Width());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Left());
    ASPOSE_ASSERT_EQ(100.0, imageShape->get_Top());
    
    ASSERT_EQ(Aspose::Words::Drawing::WrapType::Square, imageShape->get_WrapType());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, imageShape->get_RelativeHorizontalPosition());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Margin, imageShape->get_RelativeVerticalPosition());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
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

void ExDocumentBuilderImages::InsertGif()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertImage(String)
    //ExSummary:Shows how to insert gif image to the document.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    // We can insert gif image using path or bytes array.
    // It works only if DocumentBuilder optimized to Word version 2010 or higher.
    // Note, that access to the image bytes causes conversion Gif to Png.
    System::SharedPtr<Aspose::Words::Drawing::Shape> gifImage = builder->InsertImage(get_ImageDir() + u"Graphics Interchange Format.gif");
    
    gifImage = builder->InsertImage(System::IO::File::ReadAllBytes(get_ImageDir() + u"Graphics Interchange Format.gif"));
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"InsertGif.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBuilderImages, InsertGif)
{
    s_instance->InsertGif();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
