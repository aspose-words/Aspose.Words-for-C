#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class WarningInfoCollection; } }
namespace Aspose { namespace Words { class WarningInfo; } }

namespace ApiExamples {

class ExRendering : public ApiExampleBase
{
public:

    class HandleDocumentWarnings : public Aspose::Words::IWarningCallback
    {
        typedef HandleDocumentWarnings ThisType;
        typedef Aspose::Words::IWarningCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<Aspose::Words::WarningInfoCollection> FontWarnings;
        
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
    
    
public:

    void SaveToPdfWithOutline();
    void SaveToPdfStreamOnePage();
    void SaveToPdfNoCompression();
    void PdfCustomOptions();
    void SaveAsXps();
    void SaveAsXpsBookFold();
    void SaveAsImage();
    void SaveToTiffDefault();
    void SaveToTiffCompression();
    void SaveToImageResolution();
    void SaveToEmf();
    void SaveToImageJpegQuality();
    void SaveToImagePaperColor();
    void SaveToImageStream();
    void RenderToSize();
    void Thumbnails();
    void UpdatePageLayout();
    void SetTrueTypeFontsFolder();
    void SetFontsFoldersMultipleFolders();
    void SetFontsFoldersSystemAndCustomFolder();
    void SetSpecifyFontFolder();
    void SetFontSubstitutes();
    void SetSpecifyFontFolders();
    void AddFontSubstitutes();
    void SetDefaultFontName();
    void UpdatePageLayoutWarnings();
    void EmbedFullFonts();
    void SubsetFonts();
    void DisableEmbedWindowsFonts();
    void DisableEmbedCoreFonts();
    void EncryptionPermissions();
    void SetNumeralFormat();
    
};

} // namespace ApiExamples


