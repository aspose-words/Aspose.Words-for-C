#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Loading/IResourceLoadingCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Loading { enum class ResourceLoadingAction; } } }
namespace Aspose { namespace Words { namespace Loading { class ResourceLoadingArgs; } } }
namespace Aspose { namespace Words { class Document; } }

namespace ApiExamples {

class ExDocumentBase : public ApiExampleBase
{
private:

    class ImageNameHandler : public Aspose::Words::Loading::IResourceLoadingCallback
    {
        typedef ImageNameHandler ThisType;
        typedef Aspose::Words::Loading::IResourceLoadingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        Aspose::Words::Loading::ResourceLoadingAction ResourceLoading(System::SharedPtr<Aspose::Words::Loading::ResourceLoadingArgs> args) override;
        
    };
    
    
public:

    void Constructor();
    void SetPageColor();
    void ImportNode();
    void ImportNodeCustom();
    void BackgroundShape();
    void ResourceLoadingCallback();
    
protected:

    void TestResourceLoadingCallback(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples


