#pragma once

#include <system/shared_ptr.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingArgs.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingAction.h>
#include <Aspose.Words.Cpp/Model/Loading/IResourceLoadingCallback.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Loading;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExDocumentBase : public ApiExampleBase
{
    typedef ExDocumentBase ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Allows us to load images into a document using predefined shorthands, as opposed to URIs.
    /// This will separate image loading logic from the rest of the document construction.
    /// </summary>
    class ImageNameHandler : public IResourceLoadingCallback
    {
        typedef ImageNameHandler ThisType;
        typedef IResourceLoadingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        Aspose::Words::Loading::ResourceLoadingAction ResourceLoading(System::SharedPtr<Aspose::Words::Loading::ResourceLoadingArgs> args) override;
        
    };
    
    
public:

    void Constructor();
    void SetPageColor();
    void ImportNode();
    void ImportNodeCustom();
    void BackgroundShape();
    //ExStart
    //ExFor:DocumentBase.ResourceLoadingCallback
    //ExFor:IResourceLoadingCallback
    //ExFor:IResourceLoadingCallback.ResourceLoading(ResourceLoadingArgs)
    //ExFor:ResourceLoadingAction
    //ExFor:ResourceLoadingArgs
    //ExFor:ResourceLoadingArgs.OriginalUri
    //ExFor:ResourceLoadingArgs.ResourceType
    //ExFor:ResourceLoadingArgs.SetData(Byte[])
    //ExFor:ResourceType
    //ExSummary:Shows how to customize the process of loading external resources into a document.
    void ResourceLoadingCallback();
    void ImportNodeWithResolveThemeColors();
    
protected:

    //ExEnd
    void TestResourceLoadingCallback(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


