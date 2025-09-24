#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/date_time.h>
#include <system/collections/list.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IImageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Progress/IDocumentSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Progress/DocumentSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Saving;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExXamlFlowSaveOptions : public ApiExampleBase
{
    typedef ExXamlFlowSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Saving progress callback. Cancel a document saving after the "MaxDuration" seconds.
    /// </summary>
    class SavingProgressCallback : public IDocumentSavingCallback
    {
        typedef SavingProgressCallback ThisType;
        typedef IDocumentSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// Ctr.
        /// </summary>
        SavingProgressCallback();
        
        /// <summary>
        /// Callback method which called during document saving.
        /// </summary>
        /// <param name="args">Saving arguments.</param>
        void Notify(System::SharedPtr<Aspose::Words::Saving::DocumentSavingArgs> args) override;
        
    private:
    
        /// <summary>
        /// Date and time when document saving is started.
        /// </summary>
        System::DateTime mSavingStartedAt;
        /// <summary>
        /// Maximum allowed duration in sec.
        /// </summary>
        static const double MaxDuration;
        
    };
    
    
private:

    /// <summary>
    /// Counts and prints filenames of images while their parent document is converted to flow-form .xaml.
    /// </summary>
    class ImageUriPrinter : public IImageSavingCallback
    {
        typedef ImageUriPrinter ThisType;
        typedef IImageSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String get_ImagesFolderAlias() const;
        System::SharedPtr<System::Collections::Generic::List<System::String>> get_Resources() const;
        
        ImageUriPrinter(System::String imagesFolderAlias);
        
    private:
    
        System::String pr_ImagesFolderAlias;
        System::SharedPtr<System::Collections::Generic::List<System::String>> pr_Resources;
        
        void ImageSaving(System::SharedPtr<Aspose::Words::Saving::ImageSavingArgs> args) override;
        
    };
    
    
public:

    void ImageFolder();
    void ProgressCallback(Aspose::Words::SaveFormat saveFormat, System::String ext);
    void XamlReplaceBackslashWithYenSign();
    
protected:

    void TestImageFolder(System::SharedPtr<Aspose::Words::ApiExamples::ExXamlFlowSaveOptions::ImageUriPrinter> callback);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


