#pragma once

#include <system/string.h>
#include <system/collections/list.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Saving/ResourceSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IResourceSavingCallback.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Saving;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExXamlFixedSaveOptions : public ApiExampleBase
{
    typedef ExXamlFixedSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Counts and prints URIs of resources created during conversion to fixed .xaml.
    /// </summary>
    class ResourceUriPrinter : public IResourceSavingCallback
    {
        typedef ResourceUriPrinter ThisType;
        typedef IResourceSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<System::Collections::Generic::List<System::String>> get_Resources() const;
        
        ResourceUriPrinter();
        
    private:
    
        System::SharedPtr<System::Collections::Generic::List<System::String>> pr_Resources;
        
        void ResourceSaving(System::SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args) override;
        
    };
    
    
public:

    //ExStart
    //ExFor:XamlFixedSaveOptions
    //ExFor:XamlFixedSaveOptions.ResourceSavingCallback
    //ExFor:XamlFixedSaveOptions.ResourcesFolder
    //ExFor:XamlFixedSaveOptions.ResourcesFolderAlias
    //ExFor:XamlFixedSaveOptions.SaveFormat
    //ExSummary:Shows how to print the URIs of linked resources created while converting a document to fixed-form .xaml.
    void ResourceFolder();
    
protected:

    //ExEnd
    void TestResourceFolder(System::SharedPtr<Aspose::Words::ApiExamples::ExXamlFixedSaveOptions::ResourceUriPrinter> callback);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


