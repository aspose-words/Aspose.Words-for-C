#pragma once

#include <system/shared_ptr.h>
#include <system/collections/list.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Loading/BlockImportMode.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Loading;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExHtmlLoadOptions : public ApiExampleBase
{
    typedef ExHtmlLoadOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Stores all warnings that occur during a document loading operation in a List.
    /// </summary>
    class ListDocumentWarnings : public IWarningCallback
    {
        typedef ListDocumentWarnings ThisType;
        typedef IWarningCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void Warning(System::SharedPtr<Aspose::Words::WarningInfo> info) override;
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> Warnings();
        
        ListDocumentWarnings();
        
    private:
    
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> mWarnings;
        
    };
    
    
public:

    void SupportVml(bool supportVml);
    //ExStart
    //ExFor:HtmlLoadOptions.WebRequestTimeout
    //ExSummary:Shows how to set a time limit for web requests when loading a document with external resources linked by URLs.
    void WebRequestTimeout();
    //ExEnd
    void LoadHtmlFixed();
    void EncryptedHtml();
    void BaseUri();
    void GetSelectAsSdt();
    void GetInputAsFormField();
    void IgnoreNoscriptElements(bool ignoreNoscriptElements);
    void BlockImport(Aspose::Words::Loading::BlockImportMode blockImportMode);
    void FontFaceRules();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


