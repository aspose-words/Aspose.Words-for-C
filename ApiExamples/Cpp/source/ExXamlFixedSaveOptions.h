#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/collections/list.h>
#include <Aspose.Words.Cpp/Model/Saving/IResourceSavingCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Saving { class ResourceSavingArgs; } } }

namespace ApiExamples {

class ExXamlFixedSaveOptions : public ApiExampleBase
{
private:

    /// <summary>
    /// Counts and prints URIs of resources created during conversion to to fixed .xaml.
    /// </summary>
    class ResourceUriPrinter : public Aspose::Words::Saving::IResourceSavingCallback
    {
        typedef ResourceUriPrinter ThisType;
        typedef Aspose::Words::Saving::IResourceSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<System::Collections::Generic::List<System::String>> get_Resources();
        
        ResourceUriPrinter();
        
        void ResourceSaving(System::SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Collections::Generic::List<System::String>> pr_Resources;
        
    };
    
    
public:

    void ResourceFolder();
    
protected:

    void TestResourceFolder(System::SharedPtr<ExXamlFixedSaveOptions::ResourceUriPrinter> callback);
    
};

} // namespace ApiExamples


