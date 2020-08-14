#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: Aspose.Pdf is not supported

#include <system/test_tools/method_argument_tuple.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class WarningInfoCollection; } }
namespace Aspose { namespace Words { class WarningInfo; } }
namespace Aspose { namespace Words { namespace Saving { enum class HeaderFooterBookmarksExportMode; } } }
namespace ApiExamples { class ExPdfSaveOptions; }
namespace Aspose { namespace Words { namespace Saving { enum class PdfPageMode; } } }
namespace Aspose { namespace Words { namespace Saving { enum class PdfCustomPropertiesExport; } } }
namespace Aspose { namespace Words { namespace Saving { enum class DmlEffectsRenderingMode; } } }
namespace Aspose { namespace Words { namespace Saving { enum class DmlRenderingMode; } } }
namespace Aspose { namespace Words { namespace Saving { enum class EmfPlusDualRenderingMode; } } }

namespace ApiExamples {

class ExPdfSaveOptions : public ApiExampleBase
{
public:

    class HandleDocumentWarnings : public Aspose::Words::IWarningCallback
    {
        typedef HandleDocumentWarnings ThisType;
        typedef Aspose::Words::IWarningCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<Aspose::Words::WarningInfoCollection> Warnings;
        
        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        void Warning(System::SharedPtr<Aspose::Words::WarningInfo> info) override;
        
        HandleDocumentWarnings();
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    };
    
    class SaveWarningCallback : public Aspose::Words::IWarningCallback
    {
        typedef SaveWarningCallback ThisType;
        typedef Aspose::Words::IWarningCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
        friend class ExPdfSaveOptions;
        
    public:
    
        void Warning(System::SharedPtr<Aspose::Words::WarningInfo> info) override;
        
        SaveWarningCallback();
        
    protected:
    
        System::SharedPtr<Aspose::Words::WarningInfoCollection> SaveWarnings;
        
        System::Object::shared_members_type GetSharedMembers() override;
        
    };
    
    
public:

    void CreateMissingOutlineLevels();
    void TableHeadingOutlines();
    void UpdateFields(bool doUpdateFields);
    void ImageCompression();
    void ColorRendering();
    void WindowsBarPdfTitle();
    void MemoryOptimization();
    void EscapeUri(System::String uri, System::String result, bool isEscaped);
    void HandleBinaryRasterWarnings();
    void HeaderFooterBookmarksExportMode(Aspose::Words::Saving::HeaderFooterBookmarksExportMode headerFooterBookmarksExportMode);
    void UnsupportedImageFormatWarning();
    void FontsScaledToMetafileSize(bool doScaleWmfFonts);
    void AdditionalTextPositioning(bool applyAdditionalTextPositioning);
    void SaveAsPdfBookFold(bool doRenderTextAsBookfold);
    void ZoomBehaviour();
    void PageMode(Aspose::Words::Saving::PdfPageMode pageMode);
    void NoteHyperlinks(bool doCreateHyperlinks);
    void CustomPropertiesExport(Aspose::Words::Saving::PdfCustomPropertiesExport pdfCustomPropertiesExportMode);
    void DrawingMLEffects(Aspose::Words::Saving::DmlEffectsRenderingMode effectsRenderingMode);
    void DrawingMLFallback(Aspose::Words::Saving::DmlRenderingMode dmlRenderingMode);
    void ExportDocumentStructure(bool doExportStructure);
    void PreblendImages(bool doPreblendImages);
    void PdfDigitalSignature();
    void PdfDigitalSignatureTimestamp();
    void RenderMetafile(Aspose::Words::Saving::EmfPlusDualRenderingMode renderingMode);
    
};

} // namespace ApiExamples


