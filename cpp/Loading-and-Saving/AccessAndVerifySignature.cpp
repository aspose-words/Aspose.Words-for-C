#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/enumerator_adapter.h>
#include <system/date_time.h>
#include <security/cryptography/x509_certificates/x509_certificate_2.h>
#include <security/cryptography/x509_certificates/x500_distinguished_name.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatures/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatures/DigitalSignature.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatures/CertificateHolder.h>

using namespace Aspose::Words;

void AccessAndVerifySignature()
{
    // ExStart:AccessAndVerifySignature
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Test File (doc).doc");
    for (auto signature : System::IterateOver(doc->get_DigitalSignatures()))
    {
        std::cout << "*** Signature Found ***" << std::endl;
        std::cout << "Is valid: " << signature->get_IsValid() << std::endl;
        std::cout << "Reason for signing: " << signature->get_Comments().ToUtf8String() << std::endl;
        // This property is available in MS Word documents only.
        std::cout << "Time of signing: " << signature->get_SignTime().ToString().ToUtf8String() << std::endl;
        std::cout << "Subject name: " << signature->get_CertificateHolder()->get_Certificate()->get_SubjectName()->get_Name().ToUtf8String() << std::endl;
        std::cout << "Issuer name: " << signature->get_CertificateHolder()->get_Certificate()->get_IssuerName()->get_Name().ToUtf8String() << std::endl;
    }
    // ExEnd:AccessAndVerifySignature
}