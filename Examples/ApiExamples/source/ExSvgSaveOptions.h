#pragma once

#include <system/shared_ptr.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/ResourceSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IResourceSavingCallback.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Saving;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExSvgSaveOptions : public ApiExampleBase
{
    typedef ExSvgSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to .svg.
    /// </summary>
    class ResourceUriPrinter : public IResourceSavingCallback
    {
        typedef ResourceUriPrinter ThisType;
        typedef IResourceSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        ResourceUriPrinter();
        
    private:
    
        int32_t mSavedResourceCount;
        
        void ResourceSaving(System::SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args) override;
        
    };
    
    
public:

    void SaveLikeImage();
    //ExStart
    //ExFor:SvgSaveOptions
    //ExFor:SvgSaveOptions.ExportEmbeddedImages
    //ExFor:SvgSaveOptions.ResourceSavingCallback
    //ExFor:SvgSaveOptions.ResourcesFolder
    //ExFor:SvgSaveOptions.ResourcesFolderAlias
    //ExFor:SvgSaveOptions.SaveFormat
    //ExSummary:Shows how to manipulate and print the URIs of linked resources created while converting a document to .svg.
    void SvgResourceFolder();
    //ExEnd
    void SaveOfficeMath();
    void MaxImageResolution();
    void IdPrefixSvg();
    void RemoveJavaScriptFromLinksSvg();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


