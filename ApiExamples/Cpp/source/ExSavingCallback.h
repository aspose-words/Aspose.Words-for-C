#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/enum_helpers.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/IPageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/IImageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/IDocumentPartSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/ICssSavingCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Saving { class PageSavingArgs; } } }
namespace Aspose { namespace Words { namespace Saving { enum class DocumentSplitCriteria; } } }
namespace Aspose { namespace Words { namespace Saving { class DocumentPartSavingArgs; } } }
namespace Aspose { namespace Words { namespace Saving { class ImageSavingArgs; } } }
namespace Aspose { namespace Words { namespace Saving { class CssSavingArgs; } } }

namespace ApiExamples {

class ExSavingCallback : public ApiExampleBase
{
public:

    /// <summary>
    /// Renames saved images that are produced when an HTML document is saved.
    /// </summary>
    class SavedImageRename : public Aspose::Words::Saving::IImageSavingCallback
    {
        typedef SavedImageRename ThisType;
        typedef Aspose::Words::Saving::IImageSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        SavedImageRename(System::String outFileName);
        
        void ImageSaving(System::SharedPtr<Aspose::Words::Saving::ImageSavingArgs> args) override;
        
    private:
    
        int32_t mCount;
        System::String mOutFileName;
        
    };
    
    
private:

    /// <summary>
    /// Custom PageFileName is specified.
    /// </summary>
    class CustomPageFileNamePageSavingCallback : public Aspose::Words::Saving::IPageSavingCallback
    {
        typedef CustomPageFileNamePageSavingCallback ThisType;
        typedef Aspose::Words::Saving::IPageSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void PageSaving(System::SharedPtr<Aspose::Words::Saving::PageSavingArgs> args) override;
        
    };
    
    /// <summary>
    /// Renames saved document parts that are produced when an HTML document is saved while being split according to a criteria.
    /// </summary>
    class SavedDocumentPartRename : public Aspose::Words::Saving::IDocumentPartSavingCallback
    {
        typedef SavedDocumentPartRename ThisType;
        typedef Aspose::Words::Saving::IDocumentPartSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        SavedDocumentPartRename(System::String outFileName, Aspose::Words::Saving::DocumentSplitCriteria documentSplitCriteria);
        
        void DocumentPartSaving(System::SharedPtr<Aspose::Words::Saving::DocumentPartSavingArgs> args) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        int32_t mCount;
        System::String mOutFileName;
        Aspose::Words::Saving::DocumentSplitCriteria mDocumentSplitCriteria;
        
    };
    
    /// <summary>
    /// Designates a filename and other parameters for the saving of a CSS stylesheet
    /// </summary>
    class CustomCssSavingCallback : public Aspose::Words::Saving::ICssSavingCallback
    {
        typedef CustomCssSavingCallback ThisType;
        typedef Aspose::Words::Saving::ICssSavingCallback BaseType;
        
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
    void PageFileName();
    void DocumentParts();
    void CssSavingCallback();
    
};

} // namespace ApiExamples


