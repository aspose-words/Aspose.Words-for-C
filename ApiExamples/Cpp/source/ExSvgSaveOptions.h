#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/IResourceSavingCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Saving { class ResourceSavingArgs; } } }

namespace ApiExamples {

class ExSvgSaveOptions : public ApiExampleBase
{
private:

    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to .svg.
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

    void SaveLikeImage();
    void SvgResourceFolder();
    
};

} // namespace ApiExamples


