#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Document/ProtectionType.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

namespace
{

void Protect(const System::String& inputFileName)
{
    // ExStart:ProtectDocument
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputFileName);
    doc->Protect(Aspose::Words::ProtectionType::AllowOnlyFormFields, u"password");
    // ExEnd:ProtectDocument
    std::cout << "\nDocument protected successfully.\n";
}

void UnProtect(const System::String& inputFileName)
{
    // ExStart:UnProtectDocument
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputFileName);
    doc->Unprotect();
    // ExEnd:UnProtectDocument
    std::cout << "\nDocument unprotected successfully.\n";
}

void GetProtectionType(const System::String& inputFileName)
{
    // ExStart:GetProtectionType
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputFileName);
    ProtectionType protectionType = doc->get_ProtectionType();
    // ExEnd:GetProtectionType
    std::cout << "\nDocument protection type is " << static_cast<int32_t>(protectionType) << '\n';
}
}

void ProtectDocument()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument() + u"ProtectDocument.doc";
    
    Protect(dataDir);
    UnProtect(dataDir);
    GetProtectionType(dataDir);
}
