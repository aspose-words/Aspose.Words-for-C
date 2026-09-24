#pragma once

#include <system/string.h>
#include <system/date_time.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IImageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/IFontSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlVersion.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlOfficeMathOutputMode.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlMetafileFormat.h>
#include <Aspose.Words.Cpp/Model/Saving/FontSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/ExportListLabels.h>
#include <Aspose.Words.Cpp/Model/Progress/IDocumentSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Progress/DocumentSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Saving;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExHtmlSaveOptions : public ApiExampleBase
{
    typedef ExHtmlSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Prints information about exported fonts and saves them in the same local system folder as their output .html.
    /// </summary>
    class HandleFontSaving : public IFontSavingCallback
    {
        typedef HandleFontSaving ThisType;
        typedef IFontSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    private:
    
        void FontSaving(System::SharedPtr<Aspose::Words::Saving::FontSavingArgs> args) override;
        
    };
    
    /// <summary>
    /// Saving progress callback. Cancel a document saving after the "MaxDuration" seconds.
    /// </summary>
    class SavingProgressCallback : public IDocumentSavingCallback
    {
        typedef SavingProgressCallback ThisType;
        typedef IDocumentSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// Ctr.
        /// </summary>
        SavingProgressCallback();
        
        /// <summary>
        /// Callback method which called during document saving.
        /// </summary>
        /// <param name="args">Saving arguments.</param>
        void Notify(System::SharedPtr<Aspose::Words::Saving::DocumentSavingArgs> args) override;
        
    private:
    
        /// <summary>
        /// Date and time when document saving is started.
        /// </summary>
        System::DateTime mSavingStartedAt;
        /// <summary>
        /// Maximum allowed duration in sec.
        /// </summary>
        static const double MaxDuration;
        
    };
    
    
private:

    /// <summary>
    /// Prints the properties of each image as the saving process saves it to an image file in the local file system
    /// during the exporting of a document to HTML.
    /// </summary>
    class ImageShapePrinter : public IImageSavingCallback
    {
        typedef ImageShapePrinter ThisType;
        typedef IImageSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        ImageShapePrinter();
        
    private:
    
        int32_t mImageCount;
        
        void ImageSaving(System::SharedPtr<Aspose::Words::Saving::ImageSavingArgs> args) override;
        
    };
    
    
public:

    void ExportPageMarginsEpub(Aspose::Words::SaveFormat saveFormat);
    void ExportOfficeMathEpub(Aspose::Words::SaveFormat saveFormat, Aspose::Words::Saving::HtmlOfficeMathOutputMode outputMode);
    void ExportTextBoxAsSvgEpub(Aspose::Words::SaveFormat saveFormat, bool isTextBoxAsSvg);
    void CreateAZW3Toc();
    void CreateMobiToc();
    void ControlListLabelsExport(Aspose::Words::Saving::ExportListLabels howExportListLabels);
    void ExportUrlForLinkedImage(bool export_);
    void ExportRoundtripInformation();
    void RoundtripInformationDefaulValue();
    void ExternalResourceSavingConfig();
    void ConvertFontsAsBase64();
    void Html5Support(Aspose::Words::Saving::HtmlVersion htmlVersion);
    void ExportFonts(bool exportAsBase64);
    void ResourceFolderPriority();
    void ResourceFolderLowPriority();
    void SvgMetafileFormat();
    void PngMetafileFormat();
    void EmfOrWmfMetafileFormat();
    void CssClassNamesPrefix();
    void CssClassNamesNotValidPrefix();
    void CssClassNamesNullPrefix();
    void ContentIdScheme();
    void ResolveFontNames(bool resolveFontNames);
    void HeadingLevels();
    void NegativeIndent(bool allowNegativeIndent);
    void FolderAlias();
    //ExStart
    //ExFor:HtmlSaveOptions.ExportFontResources
    //ExFor:HtmlSaveOptions.FontSavingCallback
    //ExFor:IFontSavingCallback
    //ExFor:IFontSavingCallback.FontSaving
    //ExFor:FontSavingArgs
    //ExFor:FontSavingArgs.Bold
    //ExFor:FontSavingArgs.Document
    //ExFor:FontSavingArgs.FontFamilyName
    //ExFor:FontSavingArgs.FontFileName
    //ExFor:FontSavingArgs.FontStream
    //ExFor:FontSavingArgs.IsExportNeeded
    //ExFor:FontSavingArgs.IsSubsettingNeeded
    //ExFor:FontSavingArgs.Italic
    //ExFor:FontSavingArgs.KeepFontStreamOpen
    //ExFor:FontSavingArgs.OriginalFileName
    //ExFor:FontSavingArgs.OriginalFileSize
    //ExSummary:Shows how to define custom logic for exporting fonts when saving to HTML.
    void SaveExportedFonts();
    //ExEnd
    void HtmlVersions(Aspose::Words::Saving::HtmlVersion htmlVersion);
    void ExportXhtmlTransitional(bool showDoctypeDeclaration);
    void Doc2EpubSaveOptions();
    void ContentIdUrls(bool exportCidUrlsForMhtmlResources);
    void DropDownFormField(bool exportDropDownFormFieldAsText);
    void ExportImagesAsBase64(bool exportImagesAsBase64);
    void ExportFontsAsBase64();
    void ExportLanguageInformation(bool exportLanguageInformation);
    void List(Aspose::Words::Saving::ExportListLabels exportListLabels);
    void ExportPageMargins(bool exportPageMargins);
    void ExportPageSetup(bool exportPageSetup);
    void RelativeFontSize(bool exportRelativeFontSize);
    void ExportShape(bool exportShapesAsSvg);
    void RoundTripInformation(bool exportRoundtripInformation);
    void ExportTocPageNumbers(bool exportTocPageNumbers);
    void FontSubsetting(int32_t fontResourcesSubsettingSizeThreshold);
    void MetafileFormat(Aspose::Words::Saving::HtmlMetafileFormat htmlMetafileFormat);
    void OfficeMathOutputMode(Aspose::Words::Saving::HtmlOfficeMathOutputMode htmlOfficeMathOutputMode);
    void ImageFolder();
    //ExStart
    //ExFor:ImageSavingArgs.CurrentShape
    //ExFor:ImageSavingArgs.Document
    //ExFor:ImageSavingArgs.ImageStream
    //ExFor:ImageSavingArgs.IsImageAvailable
    //ExFor:ImageSavingArgs.KeepImageStreamOpen
    //ExSummary:Shows how to involve an image saving callback in an HTML conversion process.
    void ImageSavingCallback();
    //ExEnd
    void PrettyFormat(bool usePrettyFormat);
    void ProgressCallback(Aspose::Words::SaveFormat saveFormat, System::String ext);
    //ExEnd
    void MobiAzw3DefaultEncoding(Aspose::Words::SaveFormat saveFormat);
    void HtmlReplaceBackslashWithYenSign();
    void RemoveJavaScriptFromLinks();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


