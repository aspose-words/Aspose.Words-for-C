#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/DigitalSignatures/CertificateHolder.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignature.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/SignatureLine.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SignatureLineOptions.h>
#include <security/cryptography/x509_certificates/x500_distinguished_name.h>
#include <security/cryptography/x509_certificates/x509_certificate_2.h>
#include <system/array.h>
#include <system/date_time.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/guid.h>
#include <system/io/file.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;

namespace DocsExamples { namespace Programming_with_Documents { namespace Protect_or_Encrypt_Document {

class WorkingWithDigitalSinatures : public DocsExamplesBase
{
public:
    void SignDocument()
    {
        //ExStart:SingDocument
        SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        DigitalSignatureUtil::Sign(MyDir + u"Digitally signed.docx", ArtifactsDir + u"Document.Signed.docx", certHolder);
        //ExEnd:SingDocument
    }

    void SigningEncryptedDocument()
    {
        //ExStart:SigningEncryptedDocument
        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_DecryptionPassword(u"decryptionPassword");

        SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        DigitalSignatureUtil::Sign(MyDir + u"Digitally signed.docx", ArtifactsDir + u"Document.EncryptedDocument.docx", certHolder, signOptions);
        //ExEnd:SigningEncryptedDocument
    }

    void CreatingAndSigningNewSignatureLine()
    {
        //ExStart:CreatingAndSigningNewSignatureLine
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<SignatureLine> signatureLine = builder->InsertSignatureLine(MakeObject<SignatureLineOptions>())->get_SignatureLine();

        doc->Save(ArtifactsDir + u"SignDocuments.SignatureLine.docx");

        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_SignatureLineId(signatureLine->get_Id());
        signOptions->set_SignatureLineImage(System::IO::File::ReadAllBytes(ImagesDir + u"Enhanced Windows MetaFile.emf"));

        SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        DigitalSignatureUtil::Sign(ArtifactsDir + u"SignDocuments.SignatureLine.docx", ArtifactsDir + u"SignDocuments.NewSignatureLine.docx", certHolder,
                                   signOptions);
        //ExEnd:CreatingAndSigningNewSignatureLine
    }

    void SigningExistingSignatureLine()
    {
        //ExStart:SigningExistingSignatureLine
        auto doc = MakeObject<Document>(MyDir + u"Signature line.docx");

        SharedPtr<SignatureLine> signatureLine =
            (System::DynamicCast<Shape>(doc->get_FirstSection()->get_Body()->GetChild(NodeType::Shape, 0, true)))->get_SignatureLine();

        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_SignatureLineId(signatureLine->get_Id());
        signOptions->set_SignatureLineImage(System::IO::File::ReadAllBytes(ImagesDir + u"Enhanced Windows MetaFile.emf"));

        SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        DigitalSignatureUtil::Sign(MyDir + u"Digitally signed.docx", ArtifactsDir + u"SignDocuments.SigningExistingSignatureLine.docx", certHolder,
                                   signOptions);
        //ExEnd:SigningExistingSignatureLine
    }

    void SetSignatureProviderId()
    {
        //ExStart:SetSignatureProviderID
        auto doc = MakeObject<Document>(MyDir + u"Signature line.docx");

        SharedPtr<SignatureLine> signatureLine =
            (System::DynamicCast<Shape>(doc->get_FirstSection()->get_Body()->GetChild(NodeType::Shape, 0, true)))->get_SignatureLine();

        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_ProviderId(signatureLine->get_ProviderId());
        signOptions->set_SignatureLineId(signatureLine->get_Id());

        SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        DigitalSignatureUtil::Sign(MyDir + u"Digitally signed.docx", ArtifactsDir + u"SignDocuments.SetSignatureProviderId.docx", certHolder, signOptions);
        //ExEnd:SetSignatureProviderID
    }

    void CreateNewSignatureLineAndSetProviderId()
    {
        //ExStart:CreateNewSignatureLineAndSetProviderID
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto signatureLineOptions = MakeObject<SignatureLineOptions>();
        signatureLineOptions->set_Signer(u"vderyushev");
        signatureLineOptions->set_SignerTitle(u"QA");
        signatureLineOptions->set_Email(u"vderyushev@aspose.com");
        signatureLineOptions->set_ShowDate(true);
        signatureLineOptions->set_DefaultInstructions(false);
        signatureLineOptions->set_Instructions(u"Please sign here.");
        signatureLineOptions->set_AllowComments(true);

        SharedPtr<SignatureLine> signatureLine = builder->InsertSignatureLine(signatureLineOptions)->get_SignatureLine();
        signatureLine->set_ProviderId(System::Guid::Parse(u"CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2"));

        doc->Save(ArtifactsDir + u"SignDocuments.SignatureLineProviderId.docx");

        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_SignatureLineId(signatureLine->get_Id());
        signOptions->set_ProviderId(signatureLine->get_ProviderId());
        signOptions->set_Comments(u"Document was signed by vderyushev");
        signOptions->set_SignTime(System::DateTime::get_Now());

        SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        DigitalSignatureUtil::Sign(ArtifactsDir + u"SignDocuments.SignatureLineProviderId.docx",
                                   ArtifactsDir + u"SignDocuments.CreateNewSignatureLineAndSetProviderId.docx", certHolder, signOptions);
        //ExEnd:CreateNewSignatureLineAndSetProviderID
    }

    void AccessAndVerifySignature()
    {
        //ExStart:AccessAndVerifySignature
        auto doc = MakeObject<Document>(MyDir + u"Digitally signed.docx");

        for (auto signature : doc->get_DigitalSignatures())
        {
            std::cout << "*** Signature Found ***" << std::endl;
            std::cout << (String(u"Is valid: ") + signature->get_IsValid()) << std::endl;
            // This property is available in MS Word documents only.
            std::cout << (String(u"Reason for signing: ") + signature->get_Comments()) << std::endl;
            std::cout << (String(u"Time of signing: ") + signature->get_SignTime()) << std::endl;
            std::cout << (String(u"Subject name: ") + signature->get_CertificateHolder()->get_Certificate()->get_SubjectName()->get_Name()) << std::endl;
            std::cout << (String(u"Issuer name: ") + signature->get_CertificateHolder()->get_Certificate()->get_IssuerName()->get_Name()) << std::endl;
            std::cout << std::endl;
        }
        //ExEnd:AccessAndVerifySignature
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Protect_or_Encrypt_Document
