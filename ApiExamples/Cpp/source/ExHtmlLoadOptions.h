#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/test_tools/method_argument_tuple.h>
#include <system/collections/list.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class WarningInfo; } }

namespace ApiExamples {

class ExHtmlLoadOptions : public ApiExampleBase
{
private:

    /// <summary>
    /// Stores all warnings occuring during a document loading operation in a list.
    /// </summary>
    class ListDocumentWarnings : public Aspose::Words::IWarningCallback
    {
        typedef ListDocumentWarnings ThisType;
        typedef Aspose::Words::IWarningCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void Warning(System::SharedPtr<Aspose::Words::WarningInfo> info) override;
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> Warnings();
        
        ListDocumentWarnings();
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> mWarnings;
        
    };
    
    
public:

    void SupportVml(bool doSupportVml);
    void WebRequestTimeout();
    void EncryptedHtml();
    void BaseUri();
    void GetSelectAsSdt();
    void GetInputAsFormField();
    
};

} // namespace ApiExamples


