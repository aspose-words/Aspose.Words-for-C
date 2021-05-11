#pragma once

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/ProtectionType.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents { namespace Protect_or_Encrypt_Document {

class DocumentProtection : public DocsExamplesBase
{
public:
    void Protect()
    {
        //ExStart:ProtectDocument
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        doc->Protect(ProtectionType::AllowOnlyFormFields, u"password");
        //ExEnd:ProtectDocument
    }

    void Unprotect()
    {
        //ExStart:UnprotectDocument
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        doc->Unprotect();
        //ExEnd:UnprotectDocument
    }

    void GetProtectionType()
    {
        //ExStart:GetProtectionType
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        ProtectionType protectionType = doc->get_ProtectionType();
        //ExEnd:GetProtectionType
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Protect_or_Encrypt_Document
