#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/object.h>
#include <system/io/file.h>
#include <system/guid.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/SignatureLine.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SignatureLineOptions.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatures/CertificateHolder.h>

using namespace Aspose::Words;
using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;

namespace
{
    void SimpleDocumentSigning(System::String inputDataDir, System::String outputDataDir)
    {
        // ExStart:SimpleDocumentSigning
        System::SharedPtr<DigitalSignatures::CertificateHolder> certHolder = DigitalSignatures::CertificateHolder::Create(inputDataDir + u"signature.pfx", u"signature");
        System::String outputPath = outputDataDir + u"SigningSignatureLine.SimpleDocumentSigning_out.docx";
        DigitalSignatures::DigitalSignatureUtil::Sign(inputDataDir + u"Document.Signed.docx", outputPath, certHolder);
        
        // ExEnd:SimpleDocumentSigning
        std::cout << "Document is signed successfully. " << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
    
    void SigningEncryptedDocument(System::String inputDataDir, System::String outputDataDir)
    {
        // ExStart:SigningEncryptedDocument
        
        System::SharedPtr<DigitalSignatures::SignOptions> signOptions = System::MakeObject<DigitalSignatures::SignOptions>();
        signOptions->set_DecryptionPassword(u"decryptionPassword");
        
        System::SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(inputDataDir + u"signature.pfx", u"signature");
        System::String outputPath = outputDataDir + u"SigningSignatureLine.SigningEncryptedDocument_out.docx";
        DigitalSignatureUtil::Sign(inputDataDir + u"Document.Signed.docx", outputPath, certHolder, signOptions);
        // ExEnd:SigningEncryptedDocument
        std::cout << "Document is signed successfully. " << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
    
    void CreatingAndSigningNewSignatureLine(System::String inputDataDir, System::String outputDataDir)
    {
        // ExStart:CreatingAndSigningNewSignatureLine
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        System::SharedPtr<SignatureLine> signatureLine = builder->InsertSignatureLine(System::MakeObject<SignatureLineOptions>())->get_SignatureLine();
        doc->Save(outputDataDir + u"SigningSignatureLine.CreatingAndSigningNewSignatureLine1_out.docx");
        
        System::SharedPtr<SignOptions> signOptions = System::MakeObject<SignOptions>();
        signOptions->set_SignatureLineId(signatureLine->get_Id());
        signOptions->set_SignatureLineImage(System::IO::File::ReadAllBytes(inputDataDir + u"SignatureImage.emf"));
        
        System::SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(inputDataDir + u"signature.pfx", u"signature");
        System::String outputPath = outputDataDir + u"SigningSignatureLine.CreatingAndSigningNewSignatureLine2_out.docx";
        DigitalSignatureUtil::Sign(inputDataDir + u"Document.NewSignatureLine.docx", outputPath, certHolder, signOptions);
        // ExEnd:CreatingAndSigningNewSignatureLine
        
        std::cout << "Document is created and Signed with new SignatureLine successfully. " << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
    
    void SigningExistingSignatureLine(System::String inputDataDir, System::String outputDataDir)
    {
        // ExStart:SigningExistingSignatureLine
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.Signed.docx");
        System::SharedPtr<SignatureLine> signatureLine = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->get_FirstSection()->get_Body()->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_SignatureLine();
        
        System::SharedPtr<SignOptions> signOptions = System::MakeObject<SignOptions>();
        signOptions->set_SignatureLineId(signatureLine->get_Id());
        signOptions->set_SignatureLineImage(System::IO::File::ReadAllBytes(inputDataDir + u"SignatureImage.emf"));
        
        System::SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(inputDataDir + u"signature.pfx", u"signature");
        System::String outputPath = outputDataDir + u"SigningSignatureLine.SigningExistingSignatureLine_out.docx";
        DigitalSignatureUtil::Sign(inputDataDir + u"Document.Signed.docx", outputPath, certHolder, signOptions);
        // ExEnd:SigningExistingSignatureLine
        
        std::cout << "Document is signed with existing SignatureLine successfully. " << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
    
    void SetSignatureProviderID(System::String inputDataDir, System::String outputDataDir)
    {
        // ExStart:SetSignatureProviderID
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.Signed.docx");
        System::SharedPtr<SignatureLine> signatureLine = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->get_FirstSection()->get_Body()->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_SignatureLine();
        
        //Set signature and signature line provider ID
        System::SharedPtr<SignOptions> signOptions = System::MakeObject<SignOptions>();
        signOptions->set_ProviderId(signatureLine->get_ProviderId());
        signOptions->set_SignatureLineId(signatureLine->get_Id());
        
        System::SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(inputDataDir + u"signature.pfx", u"signature");
        System::String outputPath = outputDataDir + u"SigningSignatureLine.SetSignatureProviderID_out.docx";
        DigitalSignatureUtil::Sign(inputDataDir + u"Document.Signed.docx", outputPath, certHolder, signOptions);
        
        // ExEnd:SetSignatureProviderID
        
        std::cout << "Provider ID of signature is set successfully. " << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
    
    void CreateNewSignatureLineAndSetProviderID(System::String inputDataDir, System::String outputDataDir)
    {
        // ExStart:CreateNewSignatureLineAndSetProviderID
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.Signed.docx");
        
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        System::SharedPtr<SignatureLine> signatureLine = builder->InsertSignatureLine(System::MakeObject<SignatureLineOptions>())->get_SignatureLine();
        signatureLine->set_ProviderId(System::Guid(u"{F5AC7D23-DA04-45F5-ABCB-38CE7A982553}"));
        doc->Save(outputDataDir + u"SigningSignatureLine.CreateNewSignatureLineAndSetProviderID1_out.docx");
        
        System::SharedPtr<SignOptions> signOptions = System::MakeObject<SignOptions>();
        signOptions->set_SignatureLineId(signatureLine->get_Id());
        signOptions->set_ProviderId(signatureLine->get_ProviderId());
        
        System::SharedPtr<CertificateHolder> certHolder = CertificateHolder::Create(inputDataDir + u"signature.pfx", u"signature");
        System::String outputPath = outputDataDir + u"SigningSignatureLine.CreateNewSignatureLineAndSetProviderID2_out.docx";
        DigitalSignatureUtil::Sign(inputDataDir + u"Document.Signed.docx", outputPath, certHolder, signOptions);
        
        // ExEnd:CreateNewSignatureLineAndSetProviderID
        
        std::cout << "Create new signature line and set provider ID successfully. " << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void SigningSignatureLine()
{
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithSignature();
    System::String outputDataDir = GetOutputDataDir_WorkingWithSignature();
    
    if (!System::IO::File::Exists(inputDataDir + u"signature.pfx"))
    {
        std::cout << "Certificate file does not exist." << std::endl;
        return;
    }
    SimpleDocumentSigning(inputDataDir, outputDataDir);
    SigningEncryptedDocument(inputDataDir, outputDataDir);
    CreatingAndSigningNewSignatureLine(inputDataDir, outputDataDir);
    SigningExistingSignatureLine(inputDataDir, outputDataDir);
    SetSignatureProviderID(inputDataDir, outputDataDir);
    CreateNewSignatureLineAndSetProviderID(inputDataDir, outputDataDir);
}