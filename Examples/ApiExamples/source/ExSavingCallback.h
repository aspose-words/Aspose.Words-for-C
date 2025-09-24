#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/enum_helpers.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/PageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IPageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IImageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/IDocumentPartSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/ICssSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentSplitCriteria.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentPartSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/CssSavingArgs.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Saving;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExSavingCallback : public ApiExampleBase
{
    typedef ExSavingCallback ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Sets custom filenames for image files that an HTML conversion creates.
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
    /// Saves all pages to a file and directory specified within.
    /// </summary>
    class CustomFileNamePageSavingCallback : public IPageSavingCallback
    {
        typedef CustomFileNamePageSavingCallback ThisType;
        typedef IPageSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void PageSaving(System::SharedPtr<Aspose::Words::Saving::PageSavingArgs> args) override;
        
    };
    
    /// <summary>
    /// Sets custom filenames for output documents that the saving operation splits a document into.
    /// </summary>
    class SavedDocumentPartRename : public IDocumentPartSavingCallback
    {
        typedef SavedDocumentPartRename ThisType;
        typedef IDocumentPartSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        SavedDocumentPartRename(System::String outFileName, Aspose::Words::Saving::DocumentSplitCriteria documentSplitCriteria);
        
    private:
    
        int32_t mCount;
        System::String mOutFileName;
        Aspose::Words::Saving::DocumentSplitCriteria mDocumentSplitCriteria;
        
        void DocumentPartSaving(System::SharedPtr<Aspose::Words::Saving::DocumentPartSavingArgs> args) override;
        
    };
    
    /// <summary>
    /// Sets a custom filename, along with other parameters for an external CSS stylesheet.
    /// </summary>
    class CustomCssSavingCallback : public ICssSavingCallback
    {
        typedef CustomCssSavingCallback ThisType;
        typedef ICssSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        CustomCssSavingCallback(System::String cssDocFilename, bool isExportNeeded, bool keepCssStreamOpen);
        
        void CssSaving(System::SharedPtr<Aspose::Words::Saving::CssSavingArgs> args) override;
        
    private:
    
        System::String mCssTextFileName;
        bool mIsExportNeeded;
        bool mKeepCssStreamOpen;
        
    };
    
    
public:

    void CheckThatAllMethodsArePresent();
    void DocumentPartsFileNames();
    void ExternalCssFilenames();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


