#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/IPageSavingCallback.h>
#include <Aspose.Words.Cpp/Saving/ImageBinarizationMethod.h>
#include <Aspose.Words.Cpp/Saving/ImageColorMode.h>
#include <Aspose.Words.Cpp/Saving/ImagePixelFormat.h>
#include <Aspose.Words.Cpp/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/PageRange.h>
#include <Aspose.Words.Cpp/Saving/PageSavingArgs.h>
#include <Aspose.Words.Cpp/Saving/PageSet.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Saving/TiffCompression.h>
#include <system/array.h>
#include <system/object.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options {

class WorkingWithImageSaveOptions : public DocsExamplesBase
{
public:
    void ExposeThresholdControlForTiffBinarization()
    {
        //ExStart:ExposeThresholdControlForTiffBinarization
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
        saveOptions->set_TiffCompression(TiffCompression::Ccitt3);
        saveOptions->set_ImageColorMode(ImageColorMode::Grayscale);
        saveOptions->set_TiffBinarizationMethod(ImageBinarizationMethod::FloydSteinbergDithering);
        saveOptions->set_ThresholdForFloydSteinbergDithering(254);

        doc->Save(ArtifactsDir + u"WorkingWithImageSaveOptions.ExposeThresholdControlForTiffBinarization.tiff", saveOptions);
        //ExEnd:ExposeThresholdControlForTiffBinarization
    }

    void GetTiffPageRange()
    {
        //ExStart:GetTiffPageRange
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        //ExStart:SaveAsTIFF
        doc->Save(ArtifactsDir + u"WorkingWithImageSaveOptions.MultipageTiff.tiff");
        //ExEnd:SaveAsTIFF

        //ExStart:SaveAsTIFFUsingImageSaveOptions
        auto saveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Tiff);
        saveOptions->set_PageSet(MakeObject<PageSet>(MakeArray<SharedPtr<PageRange>>({MakeObject<PageRange>(0, 1)})));
        saveOptions->set_TiffCompression(TiffCompression::Ccitt4);
        saveOptions->set_Resolution(160);

        doc->Save(ArtifactsDir + u"WorkingWithImageSaveOptions.GetTiffPageRange.tiff", saveOptions);
        //ExEnd:SaveAsTIFFUsingImageSaveOptions
        //ExEnd:GetTiffPageRange
    }

    void Format1BppIndexed()
    {
        //ExStart:Format1BppIndexed
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Png);
        saveOptions->set_PageSet(MakeObject<PageSet>(1));
        saveOptions->set_ImageColorMode(ImageColorMode::BlackAndWhite);
        saveOptions->set_PixelFormat(ImagePixelFormat::Format1bppIndexed);

        doc->Save(ArtifactsDir + u"WorkingWithImageSaveOptions.Format1BppIndexed.Png", saveOptions);
        //ExEnd:Format1BppIndexed
    }

    void GetJpegPageRange()
    {
        //ExStart:GetJpegPageRange
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<ImageSaveOptions>(SaveFormat::Jpeg);

        // Set the "PageSet" to "0" to convert only the first page of a document.
        options->set_PageSet(MakeObject<PageSet>(0));

        // Change the image's brightness and contrast.
        // Both are on a 0-1 scale and are at 0.5 by default.
        options->set_ImageBrightness(0.3f);
        options->set_ImageContrast(0.7f);

        // Change the horizontal resolution.
        // The default value for these properties is 96.0, for a resolution of 96dpi.
        options->set_HorizontalResolution(72.f);

        doc->Save(ArtifactsDir + u"WorkingWithImageSaveOptions.GetJpegPageRange.jpeg", options);
        //ExEnd:GetJpegPageRange
    }

    //ExStart:PageSavingCallback
    static void PageSavingCallback()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto imageSaveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Png);
        imageSaveOptions->set_PageSet(MakeObject<PageSet>(MakeArray<SharedPtr<PageRange>>({MakeObject<PageRange>(0, doc->get_PageCount() - 1)})));
        imageSaveOptions->set_PageSavingCallback(MakeObject<WorkingWithImageSaveOptions::HandlePageSavingCallback>());

        doc->Save(ArtifactsDir + u"WorkingWithImageSaveOptions.PageSavingCallback.png", imageSaveOptions);
    }

    class HandlePageSavingCallback : public IPageSavingCallback
    {
    public:
        void PageSaving(SharedPtr<PageSavingArgs> args) override
        {
            args->set_PageFileName(String::Format(ArtifactsDir + u"Page_{0}.png", args->get_PageIndex()));
        }
    };
    //ExEnd:PageSavingCallback
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
