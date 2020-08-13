#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/test_tools/method_argument_tuple.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/IResourceSavingCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Saving { enum class HtmlFixedPageHorizontalAlignment; } } }
namespace Aspose { namespace Words { namespace Saving { class ResourceSavingArgs; } } }

namespace ApiExamples {

class ExHtmlFixedSaveOptions : public ApiExampleBase
{
private:

    class ResourceSavingCallback : public Aspose::Words::Saving::IResourceSavingCallback
    {
        typedef ResourceSavingCallback ThisType;
        typedef Aspose::Words::Saving::IResourceSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// Called when Aspose.Words saves an external resource to fixed page HTML or SVG.
        /// </summary>
        void ResourceSaving(System::SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args) override;
        
    };
    
    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to fixed .Html
    /// </summary>
    class ResourceUriPrinter : public Aspose::Words::Saving::IResourceSavingCallback
    {
        typedef ResourceUriPrinter ThisType;
        typedef Aspose::Words::Saving::IResourceSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void ResourceSaving(System::SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args) override;
        
        ResourceUriPrinter();
        
    private:
    
        int32_t mSavedResourceCount;
        
    };
    
    
public:

    void UseEncoding();
    void GetEncoding();
    void ExportEmbeddedCSS(bool doExportEmbeddedCss);
    void ExportEmbeddedFonts(bool doExportEmbeddedFonts);
    void ExportEmbeddedImages(bool doExportImages);
    void ExportEmbeddedSvgs(bool doExportSvgs);
    void ExportFormFields(bool doExportFormFields);
    void AddCssClassNamesPrefix();
    void HorizontalAlignment(Aspose::Words::Saving::HtmlFixedPageHorizontalAlignment pageHorizontalAlignment);
    void PageMargins();
    void PageMarginsException();
    void OptimizeGraphicsOutput();
    void UsingMachineFonts();
    void HtmlFixedResourceFolder();
    
};

} // namespace ApiExamples


