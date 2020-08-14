#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/test_tools/method_argument_tuple.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/IImageSavingCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { enum class SaveFormat; } }
namespace Aspose { namespace Words { namespace Saving { enum class HtmlOfficeMathOutputMode; } } }
namespace Aspose { namespace Words { namespace Saving { enum class ExportListLabels; } } }
namespace Aspose { namespace Words { namespace Saving { enum class HtmlVersion; } } }
namespace Aspose { namespace Words { namespace Saving { class ImageSavingArgs; } } }

namespace ApiExamples {

class ExHtmlSaveOptions : public ApiExampleBase
{
private:

    /// <summary>
    /// Prints information of all images that are about to be saved from within a document to image files
    /// </summary>
    class ImageShapePrinter : public Aspose::Words::Saving::IImageSavingCallback
    {
        typedef ImageShapePrinter ThisType;
        typedef Aspose::Words::Saving::IImageSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void ImageSaving(System::SharedPtr<Aspose::Words::Saving::ImageSavingArgs> args) override;
        
        ImageShapePrinter();
        
    private:
    
        int32_t mImageCount;
        
    };
    
    
public:

    void ExportPageMargins(Aspose::Words::SaveFormat saveFormat);
    void ExportOfficeMath(Aspose::Words::SaveFormat saveFormat, Aspose::Words::Saving::HtmlOfficeMathOutputMode outputMode);
    void ExportTextBoxAsSvg(Aspose::Words::SaveFormat saveFormat, bool isTextBoxAsSvg);
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
    void HeadingLevels();
    void NegativeIndent();
    void FolderAlias();
    void HtmlVersion();
    void EpubHeadings();
    void ContentIdUrls();
    void DropDownFormField();
    void ExportBase64();
    void ExportLanguageInformation();
    void List();
    void ExportPageMargins2();
    void ExportPageSetup();
    void RelativeFontSize();
    void ExportTextBox();
    void RoundTripInformation();
    void ExportTocPageNumbers();
    void FontSubsetting();
    void MetafileFormat();
    void OfficeMathOutputMode();
    void ScaleImageToShapeSize();
    void ImageSavingCallback();
    
};

} // namespace ApiExamples


