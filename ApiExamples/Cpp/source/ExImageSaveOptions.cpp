// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExImageSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <gtest/gtest.h>
#include <drawing/image.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/TiffCompression.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/MetafileRenderingOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/MetafileRenderingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/ImagePixelFormat.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageColorMode.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageBinarizationMethod.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples {

namespace gtest_test
{

class ExImageSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExImageSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExImageSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExImageSaveOptions> ExImageSaveOptions::s_instance;

} // namespace gtest_test

void ExImageSaveOptions::Renderer()
{
    //ExStart
    //ExFor:ImageSaveOptions.UseGdiEmfRenderer
    //ExSummary:Shows how to save metafiles directly without using GDI+ to EMF.
    auto doc = MakeObject<Document>(MyDir + u"Images.docx");

    auto saveOptions = MakeObject<ImageSaveOptions>(Aspose::Words::SaveFormat::Emf);
    saveOptions->set_UseGdiEmfRenderer(true);

    doc->Save(ArtifactsDir + u"ImageSaveOptions.Renderer.emf", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, Renderer)
{
    s_instance->Renderer();
}

} // namespace gtest_test

void ExImageSaveOptions::SaveSinglePage()
{
    //ExStart
    //ExFor:ImageSaveOptions.PageIndex
    //ExSummary:Shows how to save specific document page as image file.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    // For formats that can only save one page at a time,
    // the SaveOptions object can determine which page gets saved
    auto saveOptions = MakeObject<ImageSaveOptions>(Aspose::Words::SaveFormat::Gif);
    saveOptions->set_PageIndex(1);

    doc->Save(ArtifactsDir + u"ImageSaveOptions.SaveSinglePage.gif", saveOptions);
    //ExEnd

    TestUtil::VerifyImage(794, 1123, ArtifactsDir + u"ImageSaveOptions.SaveSinglePage.gif");
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, SaveSinglePage)
{
    s_instance->SaveSinglePage();
}

} // namespace gtest_test

void ExImageSaveOptions::WindowsMetaFile()
{
    //ExStart
    //ExFor:ImageSaveOptions.MetafileRenderingOptions
    //ExSummary:Shows how to set the rendering mode for Windows Metafiles.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Use a DocumentBuilder to insert a .wmf image into the document
    builder->InsertImage(System::Drawing::Image::FromFile(ImageDir + u"Windows MetaFile.wmf"));

    // Save the document as an image while setting different metafile rendering modes,
    // which will be applied to the image we inserted
    auto options = MakeObject<ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
    options->get_MetafileRenderingOptions()->set_RenderingMode(Aspose::Words::Saving::MetafileRenderingMode::Vector);

    doc->Save(ArtifactsDir + u"ImageSaveOptions.WindowsMetaFile.png", options);
    //ExEnd

    TestUtil::VerifyImage(816, 1056, ArtifactsDir + u"ImageSaveOptions.WindowsMetaFile.png");
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, WindowsMetaFile)
{
    s_instance->WindowsMetaFile();
}

} // namespace gtest_test

void ExImageSaveOptions::BlackAndWhite()
{
    //ExStart
    //ExFor:ImageColorMode
    //ExFor:ImagePixelFormat
    //ExFor:ImageSaveOptions.Clone
    //ExFor:ImageSaveOptions.ImageColorMode
    //ExFor:ImageSaveOptions.PixelFormat
    //ExSummary:Show how to convert document images to black and white with 1 bit per pixel
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    auto imageSaveOptions = MakeObject<ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
    imageSaveOptions->set_ImageColorMode(Aspose::Words::Saving::ImageColorMode::BlackAndWhite);
    imageSaveOptions->set_PixelFormat(Aspose::Words::Saving::ImagePixelFormat::Format1bppIndexed);

    // ImageSaveOptions instances can be cloned
    ASPOSE_ASSERT_NE(imageSaveOptions, imageSaveOptions->Clone());

    doc->Save(ArtifactsDir + u"ImageSaveOptions.BlackAndWhite.png", imageSaveOptions);
    //ExEnd

    TestUtil::VerifyImage(794, 1123, ArtifactsDir + u"ImageSaveOptions.BlackAndWhite.png");
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, SkipMono_BlackAndWhite)
{
    RecordProperty("category", "SkipMono");

    s_instance->BlackAndWhite();
}

} // namespace gtest_test

void ExImageSaveOptions::FloydSteinbergDithering()
{
    //ExStart
    //ExFor:ImageBinarizationMethod
    //ExFor:ImageSaveOptions.ThresholdForFloydSteinbergDithering
    //ExFor:ImageSaveOptions.TiffBinarizationMethod
    //ExSummary: Shows how to control the threshold for TIFF binarization in the Floyd-Steinberg method
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    auto options = MakeObject<ImageSaveOptions>(Aspose::Words::SaveFormat::Tiff);
    options->set_TiffCompression(Aspose::Words::Saving::TiffCompression::Ccitt3);
    options->set_ImageColorMode(Aspose::Words::Saving::ImageColorMode::Grayscale);
    options->set_TiffBinarizationMethod(Aspose::Words::Saving::ImageBinarizationMethod::FloydSteinbergDithering);
    options->set_ThresholdForFloydSteinbergDithering(254);

    doc->Save(ArtifactsDir + u"ImageSaveOptions.FloydSteinbergDithering.tiff", options);
    //ExEnd

}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, FloydSteinbergDithering)
{
    s_instance->FloydSteinbergDithering();
}

} // namespace gtest_test

void ExImageSaveOptions::EditImage()
{
    //ExStart
    //ExFor:ImageSaveOptions.HorizontalResolution
    //ExFor:ImageSaveOptions.ImageBrightness
    //ExFor:ImageSaveOptions.ImageContrast
    //ExFor:ImageSaveOptions.SaveFormat
    //ExFor:ImageSaveOptions.Scale
    //ExFor:ImageSaveOptions.VerticalResolution
    //ExSummary:Shows how to edit image.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    // When saving the document as an image, we can use an ImageSaveOptions object to edit various aspects of it
    auto options = MakeObject<ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
    options->set_ImageBrightness(0.3f);
    options->set_ImageContrast(0.7f);
    options->set_HorizontalResolution(72.f);
    options->set_VerticalResolution(72.f);
    options->set_Scale(96.f / 72.f);

    doc->Save(ArtifactsDir + u"ImageSaveOptions.EditImage.png", options);
    //ExEnd

    TestUtil::VerifyImage(794, 1123, ArtifactsDir + u"ImageSaveOptions.EditImage.png");
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, EditImage)
{
    s_instance->EditImage();
}

} // namespace gtest_test

} // namespace ApiExamples
