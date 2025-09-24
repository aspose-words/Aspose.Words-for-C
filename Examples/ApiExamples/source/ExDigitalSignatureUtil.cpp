// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDigitalSignatureUtil.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/io/stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/exceptions.h>
#include <system/details/dispose_guard.h>
#include <system/date_time.h>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/XmlDsigLevel.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureType.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignature.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/CertificateHolder.h>


using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Loading;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2520526834u, ::Aspose::Words::ApiExamples::ExDigitalSignatureUtil, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExDigitalSignatureUtil : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExDigitalSignatureUtil> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExDigitalSignatureUtil>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExDigitalSignatureUtil> ExDigitalSignatureUtil::s_instance;

} // namespace gtest_test

void ExDigitalSignatureUtil::Load()
{
    //ExStart
    //ExFor:DigitalSignatureUtil
    //ExFor:DigitalSignatureUtil.LoadSignatures(String)
    //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
    //ExSummary:Shows how to load signatures from a digitally signed document.
    // There are two ways of loading a signed document's collection of digital signatures using the DigitalSignatureUtil class.
    // 1 -  Load from a document from a local file system filename:
    System::SharedPtr<Aspose::Words::DigitalSignatures::DigitalSignatureCollection> digitalSignatures = Aspose::Words::DigitalSignatures::DigitalSignatureUtil::LoadSignatures(get_MyDir() + u"Digitally signed.docx");
    
    // If this collection is nonempty, then we can verify that the document is digitally signed.
    ASSERT_EQ(1, digitalSignatures->get_Count());
    
    // 2 -  Load from a document from a FileStream:
    {
        System::SharedPtr<System::IO::Stream> stream = System::MakeObject<System::IO::FileStream>(get_MyDir() + u"Digitally signed.docx", System::IO::FileMode::Open);
        digitalSignatures = Aspose::Words::DigitalSignatures::DigitalSignatureUtil::LoadSignatures(stream);
        ASSERT_EQ(1, digitalSignatures->get_Count());
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, Load)
{
    s_instance->Load();
}

} // namespace gtest_test

void ExDigitalSignatureUtil::Remove()
{
    //ExStart
    //ExFor:DigitalSignatureUtil
    //ExFor:DigitalSignatureUtil.LoadSignatures(String)
    //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
    //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
    //ExSummary:Shows how to remove digital signatures from a digitally signed document.
    // There are two ways of using the DigitalSignatureUtil class to remove digital signatures
    // from a signed document by saving an unsigned copy of it somewhere else in the local file system.
    // 1 - Determine the locations of both the signed document and the unsigned copy by filename strings:
    Aspose::Words::DigitalSignatures::DigitalSignatureUtil::RemoveAllSignatures(get_MyDir() + u"Digitally signed.docx", get_ArtifactsDir() + u"DigitalSignatureUtil.LoadAndRemove.FromString.docx");
    
    // 2 - Determine the locations of both the signed document and the unsigned copy by file streams:
    {
        System::SharedPtr<System::IO::Stream> streamIn = System::MakeObject<System::IO::FileStream>(get_MyDir() + u"Digitally signed.docx", System::IO::FileMode::Open);
        {
            System::SharedPtr<System::IO::Stream> streamOut = System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + u"DigitalSignatureUtil.LoadAndRemove.FromStream.docx", System::IO::FileMode::Create);
            Aspose::Words::DigitalSignatures::DigitalSignatureUtil::RemoveAllSignatures(streamIn, streamOut);
        }
    }
    
    // Verify that both our output documents have no digital signatures.
    ASSERT_EQ(0, Aspose::Words::DigitalSignatures::DigitalSignatureUtil::LoadSignatures(get_ArtifactsDir() + u"DigitalSignatureUtil.LoadAndRemove.FromString.docx")->get_Count());
    ASSERT_EQ(0, Aspose::Words::DigitalSignatures::DigitalSignatureUtil::LoadSignatures(get_ArtifactsDir() + u"DigitalSignatureUtil.LoadAndRemove.FromStream.docx")->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, Remove)
{
    s_instance->Remove();
}

} // namespace gtest_test

void ExDigitalSignatureUtil::RemoveSignatures()
{
    Aspose::Words::DigitalSignatures::DigitalSignatureUtil::RemoveAllSignatures(get_MyDir() + u"Digitally signed.odt", get_ArtifactsDir() + u"DigitalSignatureUtil.RemoveSignatures.odt");
    
    ASSERT_EQ(0, Aspose::Words::DigitalSignatures::DigitalSignatureUtil::LoadSignatures(get_ArtifactsDir() + u"DigitalSignatureUtil.RemoveSignatures.odt")->get_Count());
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, RemoveSignatures)
{
    s_instance->RemoveSignatures();
}

} // namespace gtest_test

void ExDigitalSignatureUtil::SignDocument()
{
    //ExStart
    //ExFor:CertificateHolder
    //ExFor:CertificateHolder.Create(String, String)
    //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, SignOptions)
    //ExFor:DigitalSignatures.SignOptions
    //ExFor:SignOptions.Comments
    //ExFor:SignOptions.SignTime
    //ExSummary:Shows how to digitally sign documents.
    // Create an X.509 certificate from a PKCS#12 store, which should contain a private key.
    System::SharedPtr<Aspose::Words::DigitalSignatures::CertificateHolder> certificateHolder = Aspose::Words::DigitalSignatures::CertificateHolder::Create(get_MyDir() + u"morzal.pfx", u"aw");
    
    // Create a comment and date which will be applied with our new digital signature.
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_Comments(u"My comment");
    signOptions->set_SignTime(System::DateTime::get_Now());
    
    // Take an unsigned document from the local file system via a file stream,
    // then create a signed copy of it determined by the filename of the output file stream.
    {
        System::SharedPtr<System::IO::Stream> streamIn = System::MakeObject<System::IO::FileStream>(get_MyDir() + u"Document.docx", System::IO::FileMode::Open);
        {
            System::SharedPtr<System::IO::Stream> streamOut = System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + u"DigitalSignatureUtil.SignDocument.docx", System::IO::FileMode::OpenOrCreate);
            Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(streamIn, streamOut, certificateHolder, signOptions);
        }
    }
    //ExEnd
    
    {
        System::SharedPtr<System::IO::Stream> stream = System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + u"DigitalSignatureUtil.SignDocument.docx", System::IO::FileMode::Open);
        System::SharedPtr<Aspose::Words::DigitalSignatures::DigitalSignatureCollection> digitalSignatures = Aspose::Words::DigitalSignatures::DigitalSignatureUtil::LoadSignatures(stream);
        ASSERT_EQ(1, digitalSignatures->get_Count());
        
        System::SharedPtr<Aspose::Words::DigitalSignatures::DigitalSignature> signature = digitalSignatures->idx_get(0);
        
        ASSERT_TRUE(signature->get_IsValid());
        ASSERT_EQ(Aspose::Words::DigitalSignatures::DigitalSignatureType::XmlDsig, signature->get_SignatureType());
        ASSERT_EQ(System::ObjectExt::ToString(signOptions->get_SignTime()), System::ObjectExt::ToString(signature->get_SignTime()));
        ASSERT_EQ(u"My comment", signature->get_Comments());
    }
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, SignDocument)
{
    s_instance->SignDocument();
}

} // namespace gtest_test

void ExDigitalSignatureUtil::DecryptionPassword()
{
    //ExStart
    //ExFor:CertificateHolder
    //ExFor:SignOptions.DecryptionPassword
    //ExFor:LoadOptions.Password
    //ExSummary:Shows how to sign encrypted document file.
    // Create an X.509 certificate from a PKCS#12 store, which should contain a private key.
    System::SharedPtr<Aspose::Words::DigitalSignatures::CertificateHolder> certificateHolder = Aspose::Words::DigitalSignatures::CertificateHolder::Create(get_MyDir() + u"morzal.pfx", u"aw");
    
    // Create a comment, date, and decryption password which will be applied with our new digital signature.
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_Comments(u"Comment");
    signOptions->set_SignTime(System::DateTime::get_Now());
    signOptions->set_DecryptionPassword(u"docPassword");
    
    // Set a local system filename for the unsigned input document, and an output filename for its new digitally signed copy.
    System::String inputFileName = get_MyDir() + u"Encrypted.docx";
    System::String outputFileName = get_ArtifactsDir() + u"DigitalSignatureUtil.DecryptionPassword.docx";
    
    Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(inputFileName, outputFileName, certificateHolder, signOptions);
    //ExEnd
    
    // Open encrypted document from a file.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>(u"docPassword");
    ASSERT_EQ(signOptions->get_DecryptionPassword(), loadOptions->get_Password());
    
    // Check that encrypted document was successfully signed.
    auto signedDoc = System::MakeObject<Aspose::Words::Document>(outputFileName, loadOptions);
    System::SharedPtr<Aspose::Words::DigitalSignatures::DigitalSignatureCollection> signatures = signedDoc->get_DigitalSignatures();
    
    ASSERT_EQ(1, signatures->get_Count());
    ASSERT_TRUE(signatures->get_IsValid());
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, DecryptionPassword)
{
    s_instance->DecryptionPassword();
}

} // namespace gtest_test

void ExDigitalSignatureUtil::SignDocumentObfuscationBug()
{
    System::SharedPtr<Aspose::Words::DigitalSignatures::CertificateHolder> ch = Aspose::Words::DigitalSignatures::CertificateHolder::Create(get_MyDir() + u"morzal.pfx", u"aw");
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Structured document tags.docx");
    System::String outputFileName = get_ArtifactsDir() + u"DigitalSignatureUtil.SignDocumentObfuscationBug.doc";
    
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_Comments(u"Comment");
    signOptions->set_SignTime(System::DateTime::get_Now());
    
    Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(doc->get_OriginalFileName(), outputFileName, ch, signOptions);
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, SignDocumentObfuscationBug)
{
    s_instance->SignDocumentObfuscationBug();
}

} // namespace gtest_test

void ExDigitalSignatureUtil::IncorrectDecryptionPassword()
{
    System::SharedPtr<Aspose::Words::DigitalSignatures::CertificateHolder> certificateHolder = Aspose::Words::DigitalSignatures::CertificateHolder::Create(get_MyDir() + u"morzal.pfx", u"aw");
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Encrypted.docx", System::MakeObject<Aspose::Words::Loading::LoadOptions>(u"docPassword"));
    System::String outputFileName = get_ArtifactsDir() + u"DigitalSignatureUtil.IncorrectDecryptionPassword.docx";
    
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_Comments(u"Comment");
    signOptions->set_SignTime(System::DateTime::get_Now());
    signOptions->set_DecryptionPassword(u"docPassword1");
    
    ASSERT_THROW(static_cast<std::function<void()>>([&doc, &outputFileName, &certificateHolder, &signOptions]() -> void
    {
        Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(doc->get_OriginalFileName(), outputFileName, certificateHolder, signOptions);
    })(), Aspose::Words::IncorrectPasswordException) << "The document password is incorrect.";
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, IncorrectDecryptionPassword)
{
    s_instance->IncorrectDecryptionPassword();
}

} // namespace gtest_test

void ExDigitalSignatureUtil::NoArgumentsForSing()
{
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_Comments(System::String::Empty);
    signOptions->set_SignTime(System::DateTime::get_Now());
    signOptions->set_DecryptionPassword(System::String::Empty);
    
    ASSERT_THROW(static_cast<std::function<void()>>([&signOptions]() -> void
    {
        Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(System::String::Empty, System::String::Empty, nullptr, signOptions);
    })(), System::ArgumentException);
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, NoArgumentsForSing)
{
    s_instance->NoArgumentsForSing();
}

} // namespace gtest_test

void ExDigitalSignatureUtil::NoCertificateForSign()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Digitally signed.docx");
    System::String outputFileName = get_ArtifactsDir() + u"DigitalSignatureUtil.NoCertificateForSign.docx";
    
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_Comments(u"Comment");
    signOptions->set_SignTime(System::DateTime::get_Now());
    signOptions->set_DecryptionPassword(u"docPassword");
    
    ASSERT_THROW(static_cast<std::function<void()>>([&doc, &outputFileName, &signOptions]() -> void
    {
        Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(doc->get_OriginalFileName(), outputFileName, nullptr, signOptions);
    })(), System::ArgumentNullException);
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, NoCertificateForSign)
{
    s_instance->NoCertificateForSign();
}

} // namespace gtest_test

void ExDigitalSignatureUtil::XmlDsig()
{
    //ExStart:XmlDsig
    //GistId:e06aa7a168b57907a5598e823a22bf0a
    //ExFor:SignOptions.XmlDsigLevel
    //ExFor:XmlDsigLevel
    //ExSummary:Shows how to sign document based on XML-DSig standard.
    System::SharedPtr<Aspose::Words::DigitalSignatures::CertificateHolder> certificateHolder = Aspose::Words::DigitalSignatures::CertificateHolder::Create(get_MyDir() + u"morzal.pfx", u"aw");
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_XmlDsigLevel(Aspose::Words::DigitalSignatures::XmlDsigLevel::XAdEsEpes);
    
    System::String inputFileName = get_MyDir() + u"Document.docx";
    System::String outputFileName = get_ArtifactsDir() + u"DigitalSignatureUtil.XmlDsig.docx";
    Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(inputFileName, outputFileName, certificateHolder, signOptions);
    //ExEnd:XmlDsig
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, XmlDsig)
{
    s_instance->XmlDsig();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
