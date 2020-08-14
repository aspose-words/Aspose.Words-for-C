#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/collections/list.h>
#include <Aspose.Words.Cpp/Model/Saving/IImageSavingCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Saving { class ImageSavingArgs; } } }

namespace ApiExamples {

class ExXamlFlowSaveOptions : public ApiExampleBase
{
private:

    /// <summary>
    /// Counts and prints filenames of images while their parent document is converted to flow-form .xaml.
    /// </summary>
    class ImageUriPrinter : public Aspose::Words::Saving::IImageSavingCallback
    {
        typedef ImageUriPrinter ThisType;
        typedef Aspose::Words::Saving::IImageSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String get_ImagesFolderAlias();
        System::SharedPtr<System::Collections::Generic::List<System::String>> get_Resources();
        
        ImageUriPrinter(System::String imagesFolderAlias);
        
        void ImageSaving(System::SharedPtr<Aspose::Words::Saving::ImageSavingArgs> args) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::String pr_ImagesFolderAlias;
        System::SharedPtr<System::Collections::Generic::List<System::String>> pr_Resources;
        
    };
    
    
public:

    void ImageFolder();
    
protected:

    void TestImageFolder(System::SharedPtr<ExXamlFlowSaveOptions::ImageUriPrinter> callback);
    
};

} // namespace ApiExamples


