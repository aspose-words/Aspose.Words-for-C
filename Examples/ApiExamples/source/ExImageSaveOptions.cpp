// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExImageSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/linq/enumerable.h>
#include <system/io/file_info.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/func.h>
#include <system/collections/list.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/size.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Saving/TiffCompression.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSet/PageSet.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSet/PageRange.h>
#include <Aspose.Words.Cpp/Model/Saving/MetafileRenderingOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/ImlRenderingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/ImagePageLayout/MultiPageLayout.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageBinarizationMethod.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

#include "TestUtil.h"


using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Drawing;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(4072380566u, ::Aspose::Words::ApiExamples::ExImageSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExImageSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExImageSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExImageSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExImageSaveOptions> ExImageSaveOptions::s_instance;

} // namespace gtest_test

void ExImageSaveOptions::OnePage()
{
    //ExStart
    //ExFor:Document.Save(String, SaveOptions)
    //ExFor:FixedPageSaveOptions
    //ExFor:ImageSaveOptions.PageSet
    //ExFor:PageSet
    //ExFor:PageSet.#ctor(Int32)
    //ExSummary:Shows how to render one page from a document to a JPEG image.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Page 1.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 2.");
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 3.");
    
    // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
    // to modify the way in which that method renders the document into an image.
    auto options = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Jpeg);
    // Set the "PageSet" to "1" to select the second page via
    // the zero-based index to start rendering the document from.
    options->set_PageSet(System::MakeObject<Aspose::Words::Saving::PageSet>(1));
    
    // When we save the document to the JPEG format, Aspose.Words only renders one page.
    // This image will contain one page starting from page two,
    // which will just be the second page of the original document.
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.OnePage.jpg", options);
    //ExEnd
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(816, 1056, get_ArtifactsDir() + u"ImageSaveOptions.OnePage.jpg");
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, OnePage)
{
    s_instance->OnePage();
}

} // namespace gtest_test

void ExImageSaveOptions::Renderer(bool useGdiEmfRenderer)
{
    //ExStart
    //ExFor:ImageSaveOptions.UseGdiEmfRenderer
    //ExSummary:Shows how to choose a renderer when converting a document to .emf.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
    builder->Writeln(u"Hello world!");
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    // When we save the document as an EMF image, we can pass a SaveOptions object to select a renderer for the image.
    // If we set the "UseGdiEmfRenderer" flag to "true", Aspose.Words will use the GDI+ renderer.
    // If we set the "UseGdiEmfRenderer" flag to "false", Aspose.Words will use its own metafile renderer.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Emf);
    saveOptions->set_UseGdiEmfRenderer(useGdiEmfRenderer);
    
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.Renderer.emf", saveOptions);
    //ExEnd
}

namespace gtest_test
{

using ExImageSaveOptions_Renderer_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExImageSaveOptions::Renderer)>::type;

struct ExImageSaveOptions_Renderer : public ExImageSaveOptions, public Aspose::Words::ApiExamples::ExImageSaveOptions, public ::testing::WithParamInterface<ExImageSaveOptions_Renderer_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExImageSaveOptions_Renderer, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Renderer(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExImageSaveOptions_Renderer, ::testing::ValuesIn(ExImageSaveOptions_Renderer::TestCases()));

} // namespace gtest_test

void ExImageSaveOptions::PageSet()
{
    //ExStart
    //ExFor:ImageSaveOptions.PageSet
    //ExSummary:Shows how to specify which page in a document to render as an image.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
    builder->Writeln(u"Hello world! This is page 1.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"This is page 2.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"This is page 3.");
    
    ASSERT_EQ(3, doc->get_PageCount());
    
    // When we save the document as an image, Aspose.Words only renders the first page by default.
    // We can pass a SaveOptions object to specify a different page to render.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Gif);
    // Render every page of the document to a separate image file.
    for (int32_t i = 1; i <= doc->get_PageCount(); i++)
    {
        saveOptions->set_PageSet(System::MakeObject<Aspose::Words::Saving::PageSet>(1));
        
        doc->Save(get_ArtifactsDir() + System::String::Format(u"ImageSaveOptions.PageIndex.Page {0}.gif", i), saveOptions);
    }
    //ExEnd
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(816, 1056, get_ArtifactsDir() + u"ImageSaveOptions.PageIndex.Page 1.gif");
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(816, 1056, get_ArtifactsDir() + u"ImageSaveOptions.PageIndex.Page 2.gif");
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(816, 1056, get_ArtifactsDir() + u"ImageSaveOptions.PageIndex.Page 3.gif");
    ASSERT_FALSE(System::IO::File::Exists(get_ArtifactsDir() + u"ImageSaveOptions.PageIndex.Page 4.gif"));
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, PageSet)
{
    s_instance->PageSet();
}

} // namespace gtest_test

void ExImageSaveOptions::WindowsMetaFile(Aspose::Words::Saving::MetafileRenderingMode metafileRenderingMode)
{
    //ExStart
    //ExFor:ImageSaveOptions.MetafileRenderingOptions
    //ExFor:MetafileRenderingOptions.UseGdiRasterOperationsEmulation
    //ExSummary:Shows how to set the rendering mode when saving documents with Windows Metafile images to other image formats. 
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertImage(get_ImageDir() + u"Windows MetaFile.wmf");
    
    // When we save the document as an image, we can pass a SaveOptions object to
    // determine how the saving operation will process Windows Metafiles in the document.
    // If we set the "RenderingMode" property to "MetafileRenderingMode.Vector",
    // or "MetafileRenderingMode.VectorWithFallback", we will render all metafiles as vector graphics.
    // If we set the "RenderingMode" property to "MetafileRenderingMode.Bitmap", we will render all metafiles as bitmaps.
    auto options = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
    options->get_MetafileRenderingOptions()->set_RenderingMode(metafileRenderingMode);
    // Aspose.Words uses GDI+ for raster operations emulation, when value is set to true.
    options->get_MetafileRenderingOptions()->set_UseGdiRasterOperationsEmulation(true);
    
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.WindowsMetaFile.png", options);
    //ExEnd
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(816, 1056, get_ArtifactsDir() + u"ImageSaveOptions.WindowsMetaFile.png");
}

namespace gtest_test
{

using ExImageSaveOptions_SkipMono_SkipMono_SkipMono_WindowsMetaFile_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExImageSaveOptions::WindowsMetaFile)>::type;

struct ExImageSaveOptions_SkipMono_SkipMono_SkipMono_WindowsMetaFile : public ExImageSaveOptions, public Aspose::Words::ApiExamples::ExImageSaveOptions, public ::testing::WithParamInterface<ExImageSaveOptions_SkipMono_SkipMono_SkipMono_WindowsMetaFile_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::MetafileRenderingMode::Vector),
            std::make_tuple(Aspose::Words::Saving::MetafileRenderingMode::Bitmap),
            std::make_tuple(Aspose::Words::Saving::MetafileRenderingMode::VectorWithFallback),
        };
    }
};

TEST_P(ExImageSaveOptions_SkipMono_SkipMono_SkipMono_WindowsMetaFile, Test)
{
    RecordProperty("category", "SkipMono");
    RecordProperty("category1", "SkipMono");
    RecordProperty("category2", "SkipMono");
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->WindowsMetaFile(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExImageSaveOptions_SkipMono_SkipMono_SkipMono_WindowsMetaFile, ::testing::ValuesIn(ExImageSaveOptions_SkipMono_SkipMono_SkipMono_WindowsMetaFile::TestCases()));

} // namespace gtest_test

void ExImageSaveOptions::PageByPage()
{
    //ExStart
    //ExFor:Document.Save(String, SaveOptions)
    //ExFor:FixedPageSaveOptions
    //ExFor:ImageSaveOptions.PageSet
    //ExFor:ImageSaveOptions.ImageSize
    //ExSummary:Shows how to render every page of a document to a separate TIFF image.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Page 1.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 2.");
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 3.");
    
    // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
    // to modify the way in which that method renders the document into an image.
    auto options = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Tiff);
    
    for (int32_t i = 0; i < doc->get_PageCount(); i++)
    {
        // Set the "PageSet" property to the number of the first page from
        // which to start rendering the document from.
        options->set_PageSet(System::MakeObject<Aspose::Words::Saving::PageSet>(i));
        // Export page at 2325x5325 pixels and 600 dpi.
        options->set_Resolution(600.0f);
        options->set_ImageSize(System::Drawing::Size(2325, 5325));
        
        doc->Save(get_ArtifactsDir() + System::String::Format(u"ImageSaveOptions.PageByPage.{0}.tiff", i + 1), options);
    }
    //ExEnd
    
    System::SharedPtr<System::Collections::Generic::List<System::String>> imageFileNames = System::IO::Directory::GetFiles(get_ArtifactsDir(), u"*.tiff")->LINQ_Where(static_cast<System::Func<System::String, bool>>(static_cast<std::function<bool(System::String item)>>([](System::String item) -> bool
    {
        return item.Contains(u"ImageSaveOptions.PageByPage.") && item.EndsWith(u".tiff");
    })))->LINQ_ToList();
    ASSERT_EQ(3, imageFileNames->get_Count());
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, SkipMono_PageByPage)
{
    RecordProperty("category", "SkipMono");
    
    s_instance->PageByPage();
}

} // namespace gtest_test

void ExImageSaveOptions::PaperColor()
{
    //ExStart
    //ExFor:ImageSaveOptions
    //ExFor:ImageSaveOptions.PaperColor
    //ExSummary:Renders a page of a Word document into an image with transparent or colored background.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Times New Roman");
    builder->get_Font()->set_Size(24);
    builder->Writeln(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
    // to modify the way in which that method renders the document into an image.
    auto imgOptions = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
    // Set the "PaperColor" property to a transparent color to apply a transparent
    // background to the document while rendering it to an image.
    imgOptions->set_PaperColor(System::Drawing::Color::get_Transparent());
    
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.PaperColor.Transparent.png", imgOptions);
    
    // Set the "PaperColor" property to an opaque color to apply that color
    // as the background of the document as we render it to an image.
    imgOptions->set_PaperColor(System::Drawing::Color::get_LightCoral());
    
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.PaperColor.LightCoral.png", imgOptions);
    //ExEnd
    
    Aspose::Words::ApiExamples::TestUtil::ImageContainsTransparency(get_ArtifactsDir() + u"ImageSaveOptions.PaperColor.Transparent.png");
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, PaperColor)
{
    s_instance->PaperColor();
}

} // namespace gtest_test

void ExImageSaveOptions::FloydSteinbergDithering()
{
    //ExStart
    //ExFor:ImageBinarizationMethod
    //ExFor:ImageSaveOptions.ThresholdForFloydSteinbergDithering
    //ExFor:ImageSaveOptions.TiffBinarizationMethod
    //ExSummary:Shows how to set the TIFF binarization error threshold when using the Floyd-Steinberg method to render a TIFF image.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
    builder->Writeln(u"Hello world!");
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    // When we save the document as a TIFF, we can pass a SaveOptions object to
    // adjust the dithering that Aspose.Words will apply when rendering this image.
    // The default value of the "ThresholdForFloydSteinbergDithering" property is 128.
    // Higher values tend to produce darker images.
    auto options = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Tiff);
    options->set_TiffCompression(Aspose::Words::Saving::TiffCompression::Ccitt3);
    options->set_TiffBinarizationMethod(Aspose::Words::Saving::ImageBinarizationMethod::FloydSteinbergDithering);
    options->set_ThresholdForFloydSteinbergDithering(240);
    
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.FloydSteinbergDithering.tiff", options);
    //ExEnd
    
    System::SharedPtr<System::Collections::Generic::List<System::String>> imageFileNames = System::IO::Directory::GetFiles(get_ArtifactsDir(), u"*.tiff")->LINQ_Where(static_cast<System::Func<System::String, bool>>(static_cast<std::function<bool(System::String item)>>([](System::String item) -> bool
    {
        return item.Contains(u"ImageSaveOptions.FloydSteinbergDithering.") && item.EndsWith(u".tiff");
    })))->LINQ_ToList();
    ASSERT_EQ(1, imageFileNames->get_Count());
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, SkipMono_FloydSteinbergDithering)
{
    RecordProperty("category", "SkipMono");
    
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
    //ExSummary:Shows how to edit the image while Aspose.Words converts a document to one.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
    builder->Writeln(u"Hello world!");
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    // When we save the document as an image, we can pass a SaveOptions object to
    // edit the image while the saving operation renders it.
    auto options = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
    // We can adjust these properties to change the image's brightness and contrast.
    // Both are on a 0-1 scale and are at 0.5 by default.
    options->set_ImageBrightness(0.3f);
    options->set_ImageContrast(0.7f);
    // We can adjust horizontal and vertical resolution with these properties.
    // This will affect the dimensions of the image.
    // The default value for these properties is 96.0, for a resolution of 96dpi.
    options->set_HorizontalResolution(72.f);
    options->set_VerticalResolution(72.f);
    // We can scale the image using this property. The default value is 1.0, for scaling of 100%.
    // We can use this property to negate any changes in image dimensions that changing the resolution would cause.
    options->set_Scale(96.f / 72.f);
    
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.EditImage.png", options);
    //ExEnd
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(817, 1057, get_ArtifactsDir() + u"ImageSaveOptions.EditImage.png");
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, EditImage)
{
    s_instance->EditImage();
}

} // namespace gtest_test

void ExImageSaveOptions::JpegQuality()
{
    //ExStart
    //ExFor:Document.Save(String, SaveOptions)
    //ExFor:FixedPageSaveOptions.JpegQuality
    //ExFor:ImageSaveOptions
    //ExFor:ImageSaveOptions.#ctor
    //ExFor:ImageSaveOptions.JpegQuality
    //ExSummary:Shows how to configure compression while saving a document as a JPEG.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
    // to modify the way in which that method renders the document into an image.
    auto imageOptions = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Jpeg);
    // Set the "JpegQuality" property to "10" to use stronger compression when rendering the document.
    // This will reduce the file size of the document, but the image will display more prominent compression artifacts.
    imageOptions->set_JpegQuality(10);
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.JpegQuality.HighCompression.jpg", imageOptions);
    
    // Set the "JpegQuality" property to "100" to use weaker compression when rending the document.
    // This will improve the quality of the image at the cost of an increased file size.
    imageOptions->set_JpegQuality(100);
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.JpegQuality.HighQuality.jpg", imageOptions);
    //ExEnd
    
    ASSERT_TRUE(System::MakeObject<System::IO::FileInfo>(get_ArtifactsDir() + u"ImageSaveOptions.JpegQuality.HighCompression.jpg")->get_Length() < 18000);
    ASSERT_TRUE(System::MakeObject<System::IO::FileInfo>(get_ArtifactsDir() + u"ImageSaveOptions.JpegQuality.HighQuality.jpg")->get_Length() < 75000);
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, JpegQuality)
{
    s_instance->JpegQuality();
}

} // namespace gtest_test

void ExImageSaveOptions::Resolution()
{
    //ExStart
    //ExFor:ImageSaveOptions
    //ExFor:ImageSaveOptions.Resolution
    //ExSummary:Shows how to specify a resolution while rendering a document to PNG.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Times New Roman");
    builder->get_Font()->set_Size(24);
    builder->Writeln(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
    // to modify the way in which that method renders the document into an image.
    auto options = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
    
    // Set the "Resolution" property to "72" to render the document in 72dpi.
    options->set_Resolution(72.0f);
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.Resolution.72dpi.png", options);
    
    // Set the "Resolution" property to "300" to render the document in 300dpi.
    options->set_Resolution(300.0f);
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.Resolution.300dpi.png", options);
    //ExEnd
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(612, 792, get_ArtifactsDir() + u"ImageSaveOptions.Resolution.72dpi.png");
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(2550, 3300, get_ArtifactsDir() + u"ImageSaveOptions.Resolution.300dpi.png");
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, Resolution)
{
    s_instance->Resolution();
}

} // namespace gtest_test

void ExImageSaveOptions::ExportVariousPageRanges()
{
    //ExStart
    //ExFor:PageSet.#ctor(PageRange[])
    //ExFor:PageRange
    //ExFor:PageRange.#ctor(int, int)
    //ExFor:ImageSaveOptions.PageSet
    //ExSummary:Shows how to extract pages based on exact page ranges.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Images.docx");
    
    auto imageOptions = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Tiff);
    auto pageSet = System::MakeObject<Aspose::Words::Saving::PageSet>(System::MakeArray<System::SharedPtr<Aspose::Words::Saving::PageRange>>({System::MakeObject<Aspose::Words::Saving::PageRange>(1, 1), System::MakeObject<Aspose::Words::Saving::PageRange>(2, 3), System::MakeObject<Aspose::Words::Saving::PageRange>(1, 3), System::MakeObject<Aspose::Words::Saving::PageRange>(2, 4), System::MakeObject<Aspose::Words::Saving::PageRange>(1, 1)}));
    
    imageOptions->set_PageSet(pageSet);
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.ExportVariousPageRanges.tiff", imageOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, ExportVariousPageRanges)
{
    s_instance->ExportVariousPageRanges();
}

} // namespace gtest_test

void ExImageSaveOptions::RenderInkObject()
{
    //ExStart
    //ExFor:SaveOptions.ImlRenderingMode
    //ExFor:ImlRenderingMode
    //ExSummary:Shows how to render Ink object.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Ink object.docx");
    
    // Set 'ImlRenderingMode.InkML' ignores fall-back shape of ink (InkML) object and renders InkML itself.
    // If the rendering result is unsatisfactory,
    // please use 'ImlRenderingMode.Fallback' to get a result similar to previous versions.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Jpeg);
    saveOptions->set_ImlRenderingMode(Aspose::Words::Saving::ImlRenderingMode::InkML);
    
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.RenderInkObject.jpeg", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, RenderInkObject)
{
    s_instance->RenderInkObject();
}

} // namespace gtest_test

void ExImageSaveOptions::GridLayout()
{
    //ExStart:GridLayout
    //GistId:70330eacdfc2e253f00a9adea8972975
    //ExFor:ImageSaveOptions.PageLayout
    //ExFor:MultiPageLayout
    //ExSummary:Shows how to save the document into JPG image with multi-page layout settings.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto options = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Jpeg);
    // Set up a grid layout with:
    // - 3 columns per row.
    // - 10pts spacing between pages (horizontal and vertical).
    options->set_PageLayout(Aspose::Words::Saving::MultiPageLayout::Grid(3, 10.0f, 10.0f));
    
    // Alternative layouts:
    // options.PageLayout = MultiPageLayout.Horizontal(10);
    // options.PageLayout = MultiPageLayout.Vertical(10);
    
    // Customize the background and border.
    options->get_PageLayout()->set_BackColor(System::Drawing::Color::get_LightGray());
    options->get_PageLayout()->set_BorderColor(System::Drawing::Color::get_Blue());
    options->get_PageLayout()->set_BorderWidth(2.0f);
    
    doc->Save(get_ArtifactsDir() + u"ImageSaveOptions.GridLayout.jpg", options);
    //ExEnd:GridLayout
}

namespace gtest_test
{

TEST_F(ExImageSaveOptions, GridLayout)
{
    s_instance->GridLayout();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
