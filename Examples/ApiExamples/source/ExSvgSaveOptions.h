#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/shared_ptr.h>
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
    void SvgResourceFolder();
    void SaveOfficeMath();
    void MaxImageResolution();
    void IdPrefixSvg();
    void RemoveJavaScriptFromLinksSvg();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


