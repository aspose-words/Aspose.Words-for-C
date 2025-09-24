#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/ResourceSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IResourceSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlFixedPageHorizontalAlignment.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Saving;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExHtmlFixedSaveOptions : public ApiExampleBase
{
    typedef ExHtmlFixedSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    class FontSavingCallback : public IResourceSavingCallback
    {
        typedef FontSavingCallback ThisType;
        typedef IResourceSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// Called when Aspose.Words saves an external resource to fixed page HTML or SVG.
        /// </summary>
        void ResourceSaving(System::SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args) override;
        System::String GetText();
        
        FontSavingCallback();
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mText;
        
    };
    
    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to fixed HTML.
    /// </summary>
    class ResourceUriPrinter : public IResourceSavingCallback
    {
        typedef ResourceUriPrinter ThisType;
        typedef IResourceSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String GetText();
        
        ResourceUriPrinter();
        
    private:
    
        int32_t mSavedResourceCount;
        System::SharedPtr<System::Text::StringBuilder> mText;
        
        void ResourceSaving(System::SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args) override;
        
    };
    
    
public:

    void UseEncoding();
    void GetEncoding();
    void ExportEmbeddedCss(bool exportEmbeddedCss);
    void ExportEmbeddedFonts(bool exportEmbeddedFonts);
    void ExportEmbeddedImages(bool exportImages);
    void ExportEmbeddedSvgs(bool exportSvgs);
    void ExportFormFields(bool exportFormFields);
    void HorizontalAlignment(Aspose::Words::Saving::HtmlFixedPageHorizontalAlignment pageHorizontalAlignment);
    void PageMargins();
    void PageMarginsException();
    void UsingMachineFonts(bool useTargetMachineFonts);
    void ResourceSavingCallback();
    void HtmlFixedResourceFolder();
    void IdPrefix();
    void RemoveJavaScriptFromLinks();
    
protected:

    void TestResourceSavingCallback(System::SharedPtr<Aspose::Words::ApiExamples::ExHtmlFixedSaveOptions::FontSavingCallback> callback);
    void TestHtmlFixedResourceFolder(System::SharedPtr<Aspose::Words::ApiExamples::ExHtmlFixedSaveOptions::ResourceUriPrinter> callback);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


