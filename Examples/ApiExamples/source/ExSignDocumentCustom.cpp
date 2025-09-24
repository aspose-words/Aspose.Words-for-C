// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSignDocumentCustom.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/predicate.h>
#include <gtest/gtest.h>
#include <functional>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Drawing/SignatureLine.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/SignatureLineOptions.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/CertificateHolder.h>

#include "TestUtil.h"


using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2361603606u, ::Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee, ThisTypeBaseTypesInfo);

System::Guid ExSignDocumentCustom::Signee::get_PersonId() const
{
    return pr_PersonId;
}

void ExSignDocumentCustom::Signee::set_PersonId(System::Guid value)
{
    pr_PersonId = value;
}

System::String ExSignDocumentCustom::Signee::get_Name() const
{
    return pr_Name;
}

void ExSignDocumentCustom::Signee::set_Name(System::String value)
{
    pr_Name = value;
}

System::String ExSignDocumentCustom::Signee::get_Position() const
{
    return pr_Position;
}

void ExSignDocumentCustom::Signee::set_Position(System::String value)
{
    pr_Position = value;
}

System::ArrayPtr<uint8_t> ExSignDocumentCustom::Signee::get_Image() const
{
    return pr_Image;
}

void ExSignDocumentCustom::Signee::set_Image(System::ArrayPtr<uint8_t> value)
{
    pr_Image = value;
}

ExSignDocumentCustom::Signee::Signee(System::Guid guid, System::String name, System::String position, System::ArrayPtr<uint8_t> image)
{
    set_PersonId(guid);
    set_Name(name);
    set_Position(position);
    set_Image(image);
}


RTTI_INFO_IMPL_HASH(1822251833u, ::Aspose::Words::ApiExamples::ExSignDocumentCustom, ThisTypeBaseTypesInfo);

System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee>>>& ExSignDocumentCustom::mSignees()
{
    static System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee>>> value;
    return value;
}

void ExSignDocumentCustom::SignDocument(System::String srcDocumentPath, System::String dstDocumentPath, System::SharedPtr<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee> signeeInfo, System::String certificatePath, System::String certificatePassword)
{
    auto document = System::MakeObject<Aspose::Words::Document>(srcDocumentPath);
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(document);
    
    // Configure and insert a signature line, an object in the document that will display a signature that we sign it with.
    auto signatureLineOptions = System::MakeObject<Aspose::Words::SignatureLineOptions>();
    signatureLineOptions->set_Signer(signeeInfo->get_Name());
    signatureLineOptions->set_SignerTitle(signeeInfo->get_Position());
    
    System::SharedPtr<Aspose::Words::Drawing::SignatureLine> signatureLine = builder->InsertSignatureLine(signatureLineOptions)->get_SignatureLine();
    signatureLine->set_Id(signeeInfo->get_PersonId());
    
    // First, we will save an unsigned version of our document.
    builder->get_Document()->Save(dstDocumentPath);
    
    System::SharedPtr<Aspose::Words::DigitalSignatures::CertificateHolder> certificateHolder = Aspose::Words::DigitalSignatures::CertificateHolder::Create(certificatePath, certificatePassword);
    
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_SignatureLineId(signeeInfo->get_PersonId());
    signOptions->set_SignatureLineImage(signeeInfo->get_Image());
    
    // Overwrite the unsigned document we saved above with a version signed using the certificate.
    Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(dstDocumentPath, dstDocumentPath, certificateHolder, signOptions);
}

void ExSignDocumentCustom::CreateSignees()
{
    System::String signImagePath = get_ImageDir() + u"Logo.jpg";
    
    mSignees() = [&]{ System::SharedPtr<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee> init_0[] = {System::MakeObject<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee>(System::Guid::NewGuid(), u"Ron Williams", u"Chief Executive Officer", Aspose::Words::ApiExamples::TestUtil::ImageToByteArray(signImagePath)), System::MakeObject<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee>(System::Guid::NewGuid(), u"Stephen Morse", u"Head of Compliance", Aspose::Words::ApiExamples::TestUtil::ImageToByteArray(signImagePath))}; auto list_0 = System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee>>>(); list_0->AddInitializer(2, init_0); return list_0; }();
}


namespace gtest_test
{

class ExSignDocumentCustom : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExSignDocumentCustom> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExSignDocumentCustom>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExSignDocumentCustom> ExSignDocumentCustom::s_instance;

} // namespace gtest_test

void ExSignDocumentCustom::Sign()
{
    System::String signeeName = u"Ron Williams";
    System::String srcDocumentPath = get_MyDir() + u"Document.docx";
    System::String dstDocumentPath = get_ArtifactsDir() + u"SignDocumentCustom.Sign.docx";
    System::String certificatePath = get_MyDir() + u"morzal.pfx";
    System::String certificatePassword = u"aw";
    
    CreateSignees();
    
    System::SharedPtr<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee> signeeInfo = mSignees()->Find(static_cast<System::Predicate<System::SharedPtr<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee>>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee> c)>>([&signeeName](System::SharedPtr<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee> c) -> bool
    {
        return c->get_Name() == signeeName;
    })));
    
    if (signeeInfo != nullptr)
    {
        SignDocument(srcDocumentPath, dstDocumentPath, signeeInfo, certificatePath, certificatePassword);
    }
    else
    {
        FAIL() << "Signee does not exist.";
    }
}

namespace gtest_test
{

TEST_F(ExSignDocumentCustom, Sign)
{
    s_instance->Sign();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
