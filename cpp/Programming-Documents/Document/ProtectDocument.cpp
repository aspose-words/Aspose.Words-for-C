#include "stdafx.h"
#include "examples.h"

#include <Model/Document/ProtectionType.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

namespace
{
    void Protect(const System::String& inputFileName)
    {
        // ExStart:ProtectDocument
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputFileName);
        doc->Protect(ProtectionType::AllowOnlyFormFields, u"password");
        // ExEnd:ProtectDocument
        std::cout << "Document protected successfully." << std::endl;
    }

    void UnProtect(const System::String& inputFileName)
    {
        // ExStart:UnProtectDocument
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputFileName);
        doc->Unprotect();
        // ExEnd:UnProtectDocument
        std::cout << "Document unprotected successfully." << std::endl;
    }

    void GetProtectionType(const System::String& inputFileName)
    {
        // ExStart:GetProtectionType
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputFileName);
        ProtectionType protectionType = doc->get_ProtectionType();
        // ExEnd:GetProtectionType
        std::cout << "Document protection type is " << System::ObjectExt::Box<ProtectionType>(protectionType)->ToString().ToUtf8String() << std::endl;
    }
}

void ProtectDocument()
{
    std::cout << "ProtectDocument example started." << std::endl;
    // The path to the document.
    System::String inputFileName = GetInputDataDir_WorkingWithDocument() + u"ProtectDocument.doc";
    Protect(inputFileName);
    UnProtect(inputFileName);
    GetProtectionType(inputFileName);
    std::cout << "ProtectDocument example finished." << std::endl << std::endl;
}
