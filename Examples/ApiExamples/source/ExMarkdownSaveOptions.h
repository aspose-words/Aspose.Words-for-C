#pragma once

#include <system/string.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/TableContentAlignment.h>
#include <Aspose.Words.Cpp/Model/Saving/ResourceSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/MarkdownListExportMode.h>
#include <Aspose.Words.Cpp/Model/Saving/MarkdownEmptyParagraphExportMode.h>
#include <Aspose.Words.Cpp/Model/Saving/IResourceSavingCallback.h>
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
    
    
private:

    /// <summary>
    /// Class implementing <see cref="IResourceSavingCallback"/>.
    /// </summary>
    class ChangeUriPath : public IResourceSavingCallback
    {
        typedef ChangeUriPath ThisType;
        typedef IResourceSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void ResourceSaving(System::SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args) override;
        
    };
    
    
public:

    void MarkdownDocumentTableContentAlignment(Aspose::Words::Saving::TableContentAlignment tableContentAlignment);
    //ExStart
    //ExFor:MarkdownSaveOptions
    //ExFor:MarkdownSaveOptions.#ctor
    //ExFor:MarkdownSaveOptions.ImageSavingCallback
    //ExFor:MarkdownSaveOptions.SaveFormat
    //ExFor:IImageSavingCallback
    //ExSummary:Shows how to rename the image name during saving into Markdown document.
    void RenameImages();
    //ExEnd
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
    void ResourceSavingCallback();
    //ExEnd:MarkdownResourceSavingCallback
    void ExportOfficeMathAsMarkItDown();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


