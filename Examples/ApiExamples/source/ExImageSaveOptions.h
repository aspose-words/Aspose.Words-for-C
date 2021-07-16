#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/ImageBinarizationMethod.h>
#include <Aspose.Words.Cpp/Saving/ImageColorMode.h>
#include <Aspose.Words.Cpp/Saving/ImagePixelFormat.h>
#include <Aspose.Words.Cpp/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/ImlRenderingMode.h>
#include <Aspose.Words.Cpp/Saving/MetafileRenderingMode.h>
#include <Aspose.Words.Cpp/Saving/MetafileRenderingOptions.h>
#include <Aspose.Words.Cpp/Saving/PageRange.h>
#include <Aspose.Words.Cpp/Saving/PageSet.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Saving/TiffCompression.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <drawing/color.h>
#include <drawing/image.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/collections/list.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/file_info.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <gtest/gtest-spi.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExImageSaveOptions : public ApiExampleBase
{
public:
    void OnePage()
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExFor:FixedPageSaveOptions
        //ExFor:ImageSaveOptions.PageSet
        //ExSummary:Shows how to render one page from a document to a JPEG image.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Page 1.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 2.");
        builder->InsertImage(ImageDir + u"Logo.jpg");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 3.");

        // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method renders the document into an image.
        auto options = MakeObject<ImageSaveOptions>(SaveFormat::Jpeg);

        // Set the "PageSet" to "1" to select the second page via
        // the zero-based index to start rendering the document from.
        options->set_PageSet(MakeObject<PageSet>(1));

        // When we save the document to the JPEG format, Aspose.Words only renders one page.
        // This image will contain one page starting from page two,
        // which will just be the second page of the original document.
        doc->Save(ArtifactsDir + u"ImageSaveOptions.OnePage.jpg", options);
        //ExEnd

        TestUtil::VerifyImage(816, 1056, ArtifactsDir + u"ImageSaveOptions.OnePage.jpg");
    }

    void Renderer(bool useGdiEmfRenderer)
    {
        //ExStart
        //ExFor:ImageSaveOptions.UseGdiEmfRenderer
        //ExSummary:Shows how to choose a renderer when converting a document to .emf.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
        builder->Writeln(u"Hello world!");
        builder->InsertImage(ImageDir + u"Logo.jpg");

        // When we save the document as an EMF image, we can pass a SaveOptions object to select a renderer for the image.
        // If we set the "UseGdiEmfRenderer" flag to "true", Aspose.Words will use the GDI+ renderer.
        // If we set the "UseGdiEmfRenderer" flag to "false", Aspose.Words will use its own metafile renderer.
        auto saveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Emf);
        saveOptions->set_UseGdiEmfRenderer(useGdiEmfRenderer);

        doc->Save(ArtifactsDir + u"ImageSaveOptions.Renderer.emf", saveOptions);

        ASSERT_GE(30000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"ImageSaveOptions.Renderer.emf")->get_Length());
        //ExEnd
    }

    void PageSet_()
    {
        //ExStart
        //ExFor:ImageSaveOptions.PageSet
        //ExSummary:Shows how to specify which page in a document to render as an image.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
        builder->Writeln(u"Hello world! This is page 1.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"This is page 2.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"This is page 3.");

        ASSERT_EQ(3, doc->get_PageCount());

        // When we save the document as an image, Aspose.Words only renders the first page by default.
        // We can pass a SaveOptions object to specify a different page to render.
        auto saveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Gif);

        // Render every page of the document to a separate image file.
        for (int i = 1; i <= doc->get_PageCount(); i++)
        {
            saveOptions->set_PageSet(MakeObject<PageSet>(1));

            doc->Save(ArtifactsDir + String::Format(u"ImageSaveOptions.PageIndex.Page {0}.gif", i), saveOptions);
        }
        //ExEnd

        TestUtil::VerifyImage(816, 1056, ArtifactsDir + u"ImageSaveOptions.PageIndex.Page 1.gif");
        TestUtil::VerifyImage(816, 1056, ArtifactsDir + u"ImageSaveOptions.PageIndex.Page 2.gif");
        TestUtil::VerifyImage(816, 1056, ArtifactsDir + u"ImageSaveOptions.PageIndex.Page 3.gif");
        ASSERT_FALSE(System::IO::File::Exists(ArtifactsDir + u"ImageSaveOptions.PageIndex.Page 4.gif"));
    }

    void WindowsMetaFile(MetafileRenderingMode metafileRenderingMode)
    {
        //ExStart
        //ExFor:ImageSaveOptions.MetafileRenderingOptions
        //ExSummary:Shows how to set the rendering mode when saving documents with Windows Metafile images to other image formats.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertImage(System::Drawing::Image::FromFile(ImageDir + u"Windows MetaFile.wmf"));

        // When we save the document as an image, we can pass a SaveOptions object to
        // determine how the saving operation will process Windows Metafiles in the document.
        // If we set the "RenderingMode" property to "MetafileRenderingMode.Vector",
        // or "MetafileRenderingMode.VectorWithFallback", we will render all metafiles as vector graphics.
        // If we set the "RenderingMode" property to "MetafileRenderingMode.Bitmap", we will render all metafiles as bitmaps.
        auto options = MakeObject<ImageSaveOptions>(SaveFormat::Png);
        options->get_MetafileRenderingOptions()->set_RenderingMode(metafileRenderingMode);

        doc->Save(ArtifactsDir + u"ImageSaveOptions.WindowsMetaFile.png", options);
        //ExEnd

        TestUtil::VerifyImage(816, 1056, ArtifactsDir + u"ImageSaveOptions.WindowsMetaFile.png");
    }

    void PageByPage()
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExFor:FixedPageSaveOptions
        //ExFor:ImageSaveOptions.PageSet
        //ExSummary:Shows how to render every page of a document to a separate TIFF image.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Page 1.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 2.");
        builder->InsertImage(ImageDir + u"Logo.jpg");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 3.");

        // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method renders the document into an image.
        auto options = MakeObject<ImageSaveOptions>(SaveFormat::Tiff);

        for (int i = 0; i < doc->get_PageCount(); i++)
        {
            // Set the "PageSet" property to the number of the first page from
            // which to start rendering the document from.
            options->set_PageSet(MakeObject<PageSet>(i));

            doc->Save(ArtifactsDir + String::Format(u"ImageSaveOptions.PageByPage.{0}.tiff", i + 1), options);
        }
        //ExEnd

        SharedPtr<System::Collections::Generic::List<String>> imageFileNames =
            System::IO::Directory::GetFiles(ArtifactsDir, u"*.tiff")
                ->LINQ_Where(static_cast<System::Func<String, bool>>(static_cast<std::function<bool(String item)>>(
                    [](String item) -> bool { return item.Contains(u"ImageSaveOptions.PageByPage.") && item.EndsWith(u".tiff"); })))
                ->LINQ_ToList();

        ASSERT_EQ(3, imageFileNames->get_Count());

        for (const auto& imageFileName : imageFileNames)
        {
            TestUtil::VerifyImage(816, 1056, imageFileName);
        }
    }

    void ColorMode(ImageColorMode imageColorMode)
    {
        //ExStart
        //ExFor:ImageColorMode
        //ExFor:ImageSaveOptions.ImageColorMode
        //ExSummary:Shows how to set a color mode when rendering documents.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
        builder->Writeln(u"Hello world!");
        builder->InsertImage(ImageDir + u"Logo.jpg");

        ASSERT_LT(20000, MakeObject<System::IO::FileInfo>(ImageDir + u"Logo.jpg")->get_Length());

        // When we save the document as an image, we can pass a SaveOptions object to
        // select a color mode for the image that the saving operation will generate.
        // If we set the "ImageColorMode" property to "ImageColorMode.BlackAndWhite",
        // the saving operation will apply grayscale color reduction while rendering the document.
        // If we set the "ImageColorMode" property to "ImageColorMode.Grayscale",
        // the saving operation will render the document into a monochrome image.
        // If we set the "ImageColorMode" property to "None", the saving operation will apply the default method
        // and preserve all the document's colors in the output image.
        auto imageSaveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Png);
        imageSaveOptions->set_ImageColorMode(imageColorMode);

        doc->Save(ArtifactsDir + u"ImageSaveOptions.ColorMode.png", imageSaveOptions);

        //ExEnd
    }

    void PaperColor()
    {
        //ExStart
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.PaperColor
        //ExSummary:Renders a page of a Word document into an image with transparent or colored background.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Times New Roman");
        builder->get_Font()->set_Size(24);
        builder->Writeln(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        builder->InsertImage(ImageDir + u"Logo.jpg");

        // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method renders the document into an image.
        auto imgOptions = MakeObject<ImageSaveOptions>(SaveFormat::Png);

        // Set the "PaperColor" property to a transparent color to apply a transparent
        // background to the document while rendering it to an image.
        imgOptions->set_PaperColor(System::Drawing::Color::get_Transparent());

        doc->Save(ArtifactsDir + u"ImageSaveOptions.PaperColor.Transparent.png", imgOptions);

        // Set the "PaperColor" property to an opaque color to apply that color
        // as the background of the document as we render it to an image.
        imgOptions->set_PaperColor(System::Drawing::Color::get_LightCoral());

        doc->Save(ArtifactsDir + u"ImageSaveOptions.PaperColor.LightCoral.png", imgOptions);
        //ExEnd

        TestUtil::ImageContainsTransparency(ArtifactsDir + u"ImageSaveOptions.PaperColor.Transparent.png");
        EXPECT_FATAL_FAILURE(TestUtil::ImageContainsTransparency(ArtifactsDir + u"ImageSaveOptions.PaperColor.LightCoral.png"),
                             "does not contain any transparency");
    }

    void PixelFormat(ImagePixelFormat imagePixelFormat)
    {
        //ExStart
        //ExFor:ImagePixelFormat
        //ExFor:ImageSaveOptions.Clone
        //ExFor:ImageSaveOptions.PixelFormat
        //ExSummary:Shows how to select a bit-per-pixel rate with which to render a document to an image.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
        builder->Writeln(u"Hello world!");
        builder->InsertImage(ImageDir + u"Logo.jpg");

        ASSERT_LT(20000, MakeObject<System::IO::FileInfo>(ImageDir + u"Logo.jpg")->get_Length());

        // When we save the document as an image, we can pass a SaveOptions object to
        // select a pixel format for the image that the saving operation will generate.
        // Various bit per pixel rates will affect the quality and file size of the generated image.
        auto imageSaveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Png);
        imageSaveOptions->set_PixelFormat(imagePixelFormat);

        // We can clone ImageSaveOptions instances.
        ASPOSE_ASSERT_NE(imageSaveOptions, imageSaveOptions->Clone());

        doc->Save(ArtifactsDir + u"ImageSaveOptions.PixelFormat.png", imageSaveOptions);

        //ExEnd
    }

    void FloydSteinbergDithering()
    {
        //ExStart
        //ExFor:ImageBinarizationMethod
        //ExFor:ImageSaveOptions.ThresholdForFloydSteinbergDithering
        //ExFor:ImageSaveOptions.TiffBinarizationMethod
        //ExSummary:Shows how to set the TIFF binarization error threshold when using the Floyd-Steinberg method to render a TIFF image.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
        builder->Writeln(u"Hello world!");
        builder->InsertImage(ImageDir + u"Logo.jpg");

        // When we save the document as a TIFF, we can pass a SaveOptions object to
        // adjust the dithering that Aspose.Words will apply when rendering this image.
        // The default value of the "ThresholdForFloydSteinbergDithering" property is 128.
        // Higher values tend to produce darker images.
        auto options = MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
        options->set_TiffCompression(TiffCompression::Ccitt3);
        options->set_TiffBinarizationMethod(ImageBinarizationMethod::FloydSteinbergDithering);
        options->set_ThresholdForFloydSteinbergDithering(240);

        doc->Save(ArtifactsDir + u"ImageSaveOptions.FloydSteinbergDithering.tiff", options);
        //ExEnd
    }

    void EditImage()
    {
        //ExStart
        //ExFor:ImageSaveOptions.HorizontalResolution
        //ExFor:ImageSaveOptions.ImageBrightness
        //ExFor:ImageSaveOptions.ImageContrast
        //ExFor:ImageSaveOptions.SaveFormat
        //ExFor:ImageSaveOptions.Scale
        //ExFor:ImageSaveOptions.VerticalResolution
        //ExSummary:Shows how to edit the image while Aspose.Words converts a document to one.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
        builder->Writeln(u"Hello world!");
        builder->InsertImage(ImageDir + u"Logo.jpg");

        // When we save the document as an image, we can pass a SaveOptions object to
        // edit the image while the saving operation renders it.
        auto options = MakeObject<ImageSaveOptions>(SaveFormat::Png);
        options->set_ImageBrightness(0.3f);
        options->set_ImageContrast(0.7f);
        options->set_HorizontalResolution(72.f);
        options->set_VerticalResolution(72.f);
        options->set_Scale(96.f / 72.f);

        doc->Save(ArtifactsDir + u"ImageSaveOptions.EditImage.png", options);
        //ExEnd

        TestUtil::VerifyImage(817, 1057, ArtifactsDir + u"ImageSaveOptions.EditImage.png");
    }

    void JpegQuality()
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExFor:FixedPageSaveOptions.JpegQuality
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.#ctor
        //ExFor:ImageSaveOptions.JpegQuality
        //ExSummary:Shows how to configure compression while saving a document as a JPEG.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertImage(ImageDir + u"Logo.jpg");

        // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method renders the document into an image.
        auto imageOptions = MakeObject<ImageSaveOptions>(SaveFormat::Jpeg);

        // Set the "JpegQuality" property to "10" to use stronger compression when rendering the document.
        // This will reduce the file size of the document, but the image will display more prominent compression artifacts.
        imageOptions->set_JpegQuality(10);

        doc->Save(ArtifactsDir + u"ImageSaveOptions.JpegQuality.HighCompression.jpg", imageOptions);

        ASSERT_GE(20000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"ImageSaveOptions.JpegQuality.HighCompression.jpg")->get_Length());

        // Set the "JpegQuality" property to "100" to use weaker compression when rending the document.
        // This will improve the quality of the image at the cost of an increased file size.
        imageOptions->set_JpegQuality(100);

        doc->Save(ArtifactsDir + u"ImageSaveOptions.JpegQuality.HighQuality.jpg", imageOptions);

        ASSERT_LT(60000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"ImageSaveOptions.JpegQuality.HighQuality.jpg")->get_Length());
        //ExEnd
    }

    void SaveToTiffDefault()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        doc->Save(ArtifactsDir + u"ImageSaveOptions.SaveToTiffDefault.tiff");
    }

    void TiffImageCompression(TiffCompression tiffCompression)
    {
        //ExStart
        //ExFor:TiffCompression
        //ExFor:ImageSaveOptions.TiffCompression
        //ExSummary:Shows how to select the compression scheme to apply to a document that we convert into a TIFF image.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertImage(ImageDir + u"Logo.jpg");

        // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method renders the document into an image.
        auto options = MakeObject<ImageSaveOptions>(SaveFormat::Tiff);

        // Set the "TiffCompression" property to "TiffCompression.None" to apply no compression while saving,
        // which may result in a very large output file.
        // Set the "TiffCompression" property to "TiffCompression.Rle" to apply RLE compression
        // Set the "TiffCompression" property to "TiffCompression.Lzw" to apply LZW compression.
        // Set the "TiffCompression" property to "TiffCompression.Ccitt3" to apply CCITT3 compression.
        // Set the "TiffCompression" property to "TiffCompression.Ccitt4" to apply CCITT4 compression.
        options->set_TiffCompression(tiffCompression);

        doc->Save(ArtifactsDir + u"ImageSaveOptions.TiffImageCompression.tiff", options);

        switch (tiffCompression)
        {
        case TiffCompression::None:
            ASSERT_LT(3000000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"ImageSaveOptions.TiffImageCompression.tiff")->get_Length());
            break;

        case TiffCompression::Rle:
            ASSERT_LT(600000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"ImageSaveOptions.TiffImageCompression.tiff")->get_Length());
            break;

        case TiffCompression::Lzw:
            ASSERT_LT(200000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"ImageSaveOptions.TiffImageCompression.tiff")->get_Length());
            break;

        case TiffCompression::Ccitt3:
            ASSERT_GE(90000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"ImageSaveOptions.TiffImageCompression.tiff")->get_Length());
            break;

        case TiffCompression::Ccitt4:
            ASSERT_GE(20000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"ImageSaveOptions.TiffImageCompression.tiff")->get_Length());
            break;
        }
        //ExEnd
    }

    void Resolution()
    {
        //ExStart
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.Resolution
        //ExSummary:Shows how to specify a resolution while rendering a document to PNG.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Times New Roman");
        builder->get_Font()->set_Size(24);
        builder->Writeln(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        builder->InsertImage(ImageDir + u"Logo.jpg");

        // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method renders the document into an image.
        auto options = MakeObject<ImageSaveOptions>(SaveFormat::Png);

        // Set the "Resolution" property to "72" to render the document in 72dpi.
        options->set_Resolution(72.0f);

        doc->Save(ArtifactsDir + u"ImageSaveOptions.Resolution.72dpi.png", options);

        ASSERT_GE(120000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"ImageSaveOptions.Resolution.72dpi.png")->get_Length());

        // Set the "Resolution" property to "300" to render the document in 300dpi.
        options->set_Resolution(300.0f);

        doc->Save(ArtifactsDir + u"ImageSaveOptions.Resolution.300dpi.png", options);

        ASSERT_LT(700000, MakeObject<System::IO::FileInfo>(ArtifactsDir + u"ImageSaveOptions.Resolution.300dpi.png")->get_Length());

        //ExEnd
    }

    void ExportVariousPageRanges()
    {
        //ExStart
        //ExFor:PageSet.#ctor(PageRange[])
        //ExFor:PageRange.#ctor(int, int)
        //ExFor:ImageSaveOptions.PageSet
        //ExSummary:Shows how to extract pages based on exact page ranges.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        auto imageOptions = MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
        auto pageSet = MakeObject<PageSet>(MakeArray<SharedPtr<PageRange>>(
            {MakeObject<PageRange>(1, 1), MakeObject<PageRange>(2, 3), MakeObject<PageRange>(1, 3), MakeObject<PageRange>(2, 4), MakeObject<PageRange>(1, 1)}));

        imageOptions->set_PageSet(pageSet);
        doc->Save(ArtifactsDir + u"ImageSaveOptions.ExportVariousPageRanges.tiff", imageOptions);
        //ExEnd
    }

    void RenderInkObject()
    {
        //ExStart
        //ExFor:SaveOptions.ImlRenderingMode
        //ExFor:ImlRenderingMode
        //ExSummary:Shows how to render Ink object.
        auto doc = MakeObject<Document>(MyDir + u"Ink object.docx");

        // Set 'ImlRenderingMode.InkML' ignores fall-back shape of ink (InkML) object and renders InkML itself.
        // If the rendering result is unsatisfactory,
        // please use 'ImlRenderingMode.Fallback' to get a result similar to previous versions.
        auto saveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Jpeg);
        saveOptions->set_ImlRenderingMode(ImlRenderingMode::InkML);

        doc->Save(ArtifactsDir + u"ImageSaveOptions.RenderInkObject.jpeg", saveOptions);
        //ExEnd
    }
};

} // namespace ApiExamples
