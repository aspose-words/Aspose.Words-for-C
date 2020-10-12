#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageBinarizationMethod.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageColorMode.h>
#include <Aspose.Words.Cpp/Model/Saving/ImagePixelFormat.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/TiffCompression.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSet/PageSet.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSet/PageRange.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    /*void SaveColorTIFFwithLZW(System::SharedPtr<Document> doc, System::String const &outputDataDir, float brightness, float contrast)
    {
        // Select the TIFF format with 100 dpi.
        System::SharedPtr<ImageSaveOptions> imgOpttiff = System::MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
        imgOpttiff->set_Resolution(100.0f);

        // Select fullcolor LZW compression.
        imgOpttiff->set_TiffCompression(TiffCompression::Lzw);

        // Set brightness and contrast.
        imgOpttiff->set_ImageBrightness(brightness);
        imgOpttiff->set_ImageContrast(contrast);

        // Save multipage color TIFF.
        System::String outputPath = outputDataDir + u"ImageColorFilters.SaveColorTIFFwithLZW.tiff";
        doc->Save(outputPath, imgOpttiff);

        std::cout << "Document converted to TIFF successfully with Colors." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }*/

    /*void SaveGrayscaleTIFFwithLZW(System::SharedPtr<Document> doc, System::String const &outputDataDir, float brightness, float contrast)
    {
        // Select the TIFF format with 100 dpi.
        System::SharedPtr<ImageSaveOptions> imgOpttiff = System::MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
        imgOpttiff->set_Resolution(100.0f);

        // Select LZW compression.
        imgOpttiff->set_TiffCompression(TiffCompression::Lzw);

        // Apply grayscale filter.
        imgOpttiff->set_ImageColorMode(ImageColorMode::Grayscale);

        // Set brightness and contrast.
        imgOpttiff->set_ImageBrightness(brightness);
        imgOpttiff->set_ImageContrast(contrast);

        // Save multipage grayscale TIFF.
        System::String outputPath = outputDataDir + u"ImageColorFilters.SaveGrayscaleTIFFwithLZW.tiff";
        doc->Save(outputPath, imgOpttiff);

        std::cout << "Document converted to TIFF successfully with Gray scale." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }*/

    /*void SaveBlackWhiteTIFFwithLZW(System::SharedPtr<Document> doc, System::String const &outputDataDir, bool highSensitivity)
    {
        // Select the TIFF format with 100 dpi.
        System::SharedPtr<ImageSaveOptions> imgOpttiff = System::MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
        imgOpttiff->set_Resolution(100.0f);

        // Apply black & white filter. Set very high sensitivity to gray color.
        imgOpttiff->set_TiffCompression(TiffCompression::Lzw);
        imgOpttiff->set_ImageColorMode(ImageColorMode::BlackAndWhite);

        // Set brightness and contrast according to sensitivity.
        if (highSensitivity)
        {
            imgOpttiff->set_ImageBrightness(0.4f);
            imgOpttiff->set_ImageContrast(0.3f);
        }
        else
        {
            imgOpttiff->set_ImageBrightness(0.9f);
            imgOpttiff->set_ImageContrast(0.9f);
        }

        // Save multipage TIFF.
        System::String outputPath = outputDataDir + u"ImageColorFilters.SaveBlackWhiteTIFFwithLZW.tiff";
        doc->Save(outputPath, imgOpttiff);

        std::cout << "Document converted to TIFF successfully with black and white." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }*/

    /*void SaveBlackWhiteTIFFwithCITT4(System::SharedPtr<Document> doc, System::String const &outputDataDir, bool highSensitivity)
    {
        // Select the TIFF format with 100 dpi.
        System::SharedPtr<ImageSaveOptions> imgOpttiff = System::MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
        imgOpttiff->set_Resolution(100.0f);

        // Set CCITT4 compression.
        imgOpttiff->set_TiffCompression(TiffCompression::Ccitt4);

        // Apply grayscale filter.
        imgOpttiff->set_ImageColorMode(ImageColorMode::Grayscale);

        // Set brightness and contrast according to sensitivity.
        if (highSensitivity)
        {
            imgOpttiff->set_ImageBrightness(0.4f);
            imgOpttiff->set_ImageContrast(0.3f);
        }
        else
        {
            imgOpttiff->set_ImageBrightness(0.9f);
            imgOpttiff->set_ImageContrast(0.9f);
        }

        // Save multipage TIFF.
        System::String outputPath = outputDataDir + u"ImageColorFilters.SaveBlackWhiteTIFFwithCITT4.tiff";
        doc->Save(outputPath, imgOpttiff);

        std::cout << "Document converted to TIFF successfully with black and white and Ccitt4 compression." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }*/

    /*void SaveBlackWhiteTIFFwithRLE(System::SharedPtr<Document> doc, System::String const &outputDataDir, bool highSensitivity)
    {
        // Select the TIFF format with 100 dpi.
        System::SharedPtr<ImageSaveOptions> imgOpttiff = System::MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
        imgOpttiff->set_Resolution(100.0f);

        // Set RLE compression.
        imgOpttiff->set_TiffCompression(TiffCompression::Rle);

        // Aply grayscale filter.
        imgOpttiff->set_ImageColorMode(ImageColorMode::Grayscale);

        // Set brightness and contrast according to sensitivity.
        if (highSensitivity)
        {
            imgOpttiff->set_ImageBrightness(0.4f);
            imgOpttiff->set_ImageContrast(0.3f);
        }
        else
        {
            imgOpttiff->set_ImageBrightness(0.9f);
            imgOpttiff->set_ImageContrast(0.9f);
        }

        // Save multipage TIFF grayscale with low bright and contrast
        System::String outputPath = outputDataDir + u"ImageColorFilters.SaveBlackWhiteTIFFwithRLE.tiff";
        doc->Save(outputPath, imgOpttiff);

        std::cout << "Document converted to TIFF successfully with black and white and Rle compression." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }*/

    void SaveImageToOnebitPerPixel(System::SharedPtr<Document> doc, System::String const &outputDataDir)
    {
        // ExStart:SaveImageToOnebitPerPixel
        System::SharedPtr<ImageSaveOptions> opt = System::MakeObject<ImageSaveOptions>(SaveFormat::Png);

        auto pageSet = System::MakeObject<PageSet>(System::MakeArray<System::SharedPtr<PageRange>>({System::MakeObject<PageRange>(1,1)}));
        opt->set_PageSet(pageSet);

        opt->set_ImageColorMode(ImageColorMode::BlackAndWhite);
        opt->set_PixelFormat(ImagePixelFormat::Format1bppIndexed);

        System::String outputPath = outputDataDir + u"ImageColorFilters.SaveImageToOnebitPerPixel.png";
        doc->Save(outputPath, opt);
        // ExEnd:SaveImageToOnebitPerPixel
        std::cout << "Document converted to PNG successfully with 1 bit per pixel." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    /*void ExposeThresholdControlForTiffBinarization(System::SharedPtr<Document> doc, System::String const &outputDataDir)
    {
        // ExStart:ExposeThresholdControlForTiffBinarization
        System::SharedPtr<ImageSaveOptions> options = System::MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
        options->set_TiffCompression(TiffCompression::Ccitt3);
        options->set_ImageColorMode(ImageColorMode::Grayscale);
        options->set_TiffBinarizationMethod(ImageBinarizationMethod::FloydSteinbergDithering);
        options->set_ThresholdForFloydSteinbergDithering(254);

        System::String outputPath = outputDataDir + u"ImageColorFilters.ExposeThresholdControlForTiffBinarization.tiff";
        doc->Save(outputPath, options);
        // ExEnd:ExposeThresholdControlForTiffBinarization
        std::cout << "Expose Threshold Control For Tiff Binarization." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }*/
}

void ImageColorFilters()
{
    std::cout << "ImageColorFilters example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();
    // Open the document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.Colors.docx");
#if 0
    // Tiff support is limited and unstable
    SaveColorTIFFwithLZW(doc, outputDataDir, 0.8f, 0.8f); 
    SaveGrayscaleTIFFwithLZW(doc, outputDataDir, 0.8f, 0.8f); 
    aveBlackWhiteTIFFwithLZW(doc, outputDataDir, true); 
    SaveBlackWhiteTIFFwithCITT4(doc, outputDataDir, true); 
    SaveBlackWhiteTIFFwithRLE(doc, outputDataDir, true); 
    ExposeThresholdControlForTiffBinarization(doc, outputDataDir); 
#endif
    SaveImageToOnebitPerPixel(doc, outputDataDir);
    std::cout << "ImageColorFilters example finished." << std::endl << std::endl;
}