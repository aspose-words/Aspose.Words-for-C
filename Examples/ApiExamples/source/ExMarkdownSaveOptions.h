#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/TableContentAlignment.h>
#include <Aspose.Words.Cpp/Model/Saving/MarkdownListExportMode.h>
#include <Aspose.Words.Cpp/Model/Saving/MarkdownEmptyParagraphExportMode.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IImageSavingCallback.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Saving;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExMarkdownSaveOptions : public ApiExampleBase
{
    typedef ExMarkdownSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Renames saved images that are produced when an Markdown document is saved.
    /// </summary>
    class SavedImageRename : public IImageSavingCallback
    {
        typedef SavedImageRename ThisType;
        typedef IImageSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        SavedImageRename(System::String outFileName);
        
    private:
    
        int32_t mCount;
        System::String mOutFileName;
        
        void ImageSaving(System::SharedPtr<Aspose::Words::Saving::ImageSavingArgs> args) override;
        
    };
    
    
public:

    void MarkdownDocumentTableContentAlignment(Aspose::Words::Saving::TableContentAlignment tableContentAlignment);
    void RenameImages();
    void ExportImagesAsBase64(bool exportImagesAsBase64);
    void ListExportMode(Aspose::Words::Saving::MarkdownListExportMode markdownListExportMode);
    void ImagesFolder();
    void ExportUnderlineFormatting();
    void LinkExportMode();
    void ExportTableAsHtml();
    void ImageResolution();
    void OfficeMathExportMode();
    void EmptyParagraphExportMode(Aspose::Words::Saving::MarkdownEmptyParagraphExportMode exportMode);
    void ExportOfficeMathAsLatex();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


