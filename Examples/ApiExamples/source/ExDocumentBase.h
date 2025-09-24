#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/shared_ptr.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingArgs.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingAction.h>
#include <Aspose.Words.Cpp/Model/Loading/IResourceLoadingCallback.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Loading;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExDocumentBase : public ApiExampleBase
{
    typedef ExDocumentBase ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Allows us to load images into a document using predefined shorthands, as opposed to URIs.
    /// This will separate image loading logic from the rest of the document construction.
    /// </summary>
    class ImageNameHandler : public IResourceLoadingCallback
    {
        typedef ImageNameHandler ThisType;
        typedef IResourceLoadingCallback BaseType;
        
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
} // namespace Words
} // namespace Aspose


