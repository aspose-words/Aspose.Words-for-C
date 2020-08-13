// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDigitalSignatureUtil.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/io/stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/exceptions.h>
#include <system/details/dispose_guard.h>
#include <system/date_time.h>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/SignOptions.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureType.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignature.h>
#include <Aspose.Words.Cpp/Model/Document/CertificateHolder.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
namespace ApiExamples {

namespace gtest_test
{

class ExDigitalSignatureUtil : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExDigitalSignatureUtil> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExDigitalSignatureUtil>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExDigitalSignatureUtil> ExDigitalSignatureUtil::s_instance;

} // namespace gtest_test

void ExDigitalSignatureUtil::LoadAndRemove()
{
    //ExStart
    //ExFor:DigitalSignatureUtil
    //ExFor:DigitalSignatureUtil.LoadSignatures(String)
    //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
    //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
    //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
    //ExSummary:Shows how to load and remove digital signatures from a digitally signed document.
    // Load digital signatures via filename string to verify that the document is signed
    SharedPtr<DigitalSignatureCollection> digitalSignatures = DigitalSignatureUtil::LoadSignatures(MyDir + u"Digitally signed.docx");
    ASSERT_EQ(1, digitalSignatures->get_Count());

    // Re-save the document to an output filename with all digital signatures removed
    DigitalSignatureUtil::RemoveAllSignatures(MyDir + u"Digitally signed.docx", ArtifactsDir + u"DigitalSignatureUtil.LoadAndRemove.FromString.docx");

    // Remove all signatures from the document using stream parameters
    {
        SharedPtr<System::IO::Stream> streamIn = MakeObject<System::IO::FileStream>(MyDir + u"Digitally signed.docx", System::IO::FileMode::Open);
        {
            SharedPtr<System::IO::Stream> streamOut = MakeObject<System::IO::FileStream>(ArtifactsDir + u"DigitalSignatureUtil.LoadAndRemove.FromStream.docx", System::IO::FileMode::Create);
            DigitalSignatureUtil::RemoveAllSignatures(streamIn, streamOut);
        }
    }

    // We can also load a document's digital signatures via stream, which we will do to verify that all signatures have been removed
    {
        SharedPtr<System::IO::Stream> stream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"DigitalSignatureUtil.LoadAndRemove.FromStream.docx", System::IO::FileMode::Open);
        digitalSignatures = DigitalSignatureUtil::LoadSignatures(stream);
        ASSERT_EQ(0, digitalSignatures->get_Count());
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, LoadAndRemove)
{
    s_instance->LoadAndRemove();
}

} // namespace gtest_test

void ExDigitalSignatureUtil::SignDocument()
{
    //ExStart
    //ExFor:CertificateHolder
    //ExFor:CertificateHolder.Create(String, String)
    //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, SignOptions)
    //ExFor:SignOptions.Comments
    //ExFor:SignOptions.SignTime
    //ExSummary:Shows how to sign documents using certificate holder and sign options.
    SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

    auto signOptions = MakeObject<SignOptions>();
    signOptions->set_Comments(u"My comment");
    signOptions->set_SignTime(System::DateTime::get_Now());

    {
        SharedPtr<System::IO::Stream> streamIn = MakeObject<System::IO::FileStream>(MyDir + u"Digitally signed.docx", System::IO::FileMode::Open);
        {
            SharedPtr<System::IO::Stream> streamOut = MakeObject<System::IO::FileStream>(ArtifactsDir + u"DigitalSignatureUtil.SignDocument.docx", System::IO::FileMode::OpenOrCreate);
            DigitalSignatureUtil::Sign(streamIn, streamOut, certificateHolder, signOptions);
        }
    }
    //ExEnd

    {
        SharedPtr<System::IO::Stream> stream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"DigitalSignatureUtil.SignDocument.docx", System::IO::FileMode::Open);
        SharedPtr<DigitalSignatureCollection> digitalSignatures = DigitalSignatureUtil::LoadSignatures(stream);
        ASSERT_EQ(1, digitalSignatures->get_Count());

        SharedPtr<DigitalSignature> signature = digitalSignatures->idx_get(0);

        ASSERT_TRUE(signature->get_IsValid());
        ASSERT_EQ(Aspose::Words::DigitalSignatureType::XmlDsig, signature->get_SignatureType());
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

void ExDigitalSignatureUtil::SignDocumentObfuscationBug()
{
    SharedPtr<CertificateHolder> ch = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

    auto doc = MakeObject<Document>(MyDir + u"Structured document tags.docx");
    String outputFileName = ArtifactsDir + u"DigitalSignatureUtil.SignDocumentObfuscationBug.doc";

    auto signOptions = MakeObject<SignOptions>();
    signOptions->set_Comments(u"Comment");
    signOptions->set_SignTime(System::DateTime::get_Now());

    DigitalSignatureUtil::Sign(doc->get_OriginalFileName(), outputFileName, ch, signOptions);
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
    SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

    auto doc = MakeObject<Document>(MyDir + u"Encrypted.docx", MakeObject<LoadOptions>(u"docPassword"));
    String outputFileName = ArtifactsDir + u"DigitalSignatureUtil.IncorrectDecryptionPassword.docx";

    auto signOptions = MakeObject<SignOptions>();
    signOptions->set_Comments(u"Comment");
    signOptions->set_SignTime(System::DateTime::get_Now());
    signOptions->set_DecryptionPassword(u"docPassword1");

    // Digitally sign encrypted with "docPassword" document in the specified path

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_0 = [&doc, &outputFileName, &certificateHolder, &signOptions]()
    {
        DigitalSignatureUtil::Sign(doc->get_OriginalFileName(), outputFileName, certificateHolder, signOptions);
    };

    ASSERT_THROW(_local_func_0(), IncorrectPasswordException) << "The document password is incorrect.";
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, IncorrectDecryptionPassword)
{
    s_instance->IncorrectDecryptionPassword();
}

} // namespace gtest_test

void ExDigitalSignatureUtil::DecryptionPassword()
{
    //ExStart
    //ExFor:CertificateHolder
    //ExFor:SignOptions.DecryptionPassword
    //ExFor:LoadOptions.Password
    //ExSummary:Shows how to sign encrypted document file.
    // Create certificate holder from a file
    SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

    auto signOptions = MakeObject<SignOptions>();
    signOptions->set_Comments(u"Comment");
    signOptions->set_SignTime(System::DateTime::get_Now());
    signOptions->set_DecryptionPassword(u"docPassword");

    // Digitally sign encrypted with "docPassword" document in the specified path
    String inputFileName = MyDir + u"Encrypted.docx";
    String outputFileName = ArtifactsDir + u"DigitalSignatureUtil.DecryptionPassword.docx";

    DigitalSignatureUtil::Sign(inputFileName, outputFileName, certificateHolder, signOptions);
    //ExEnd

    // Open encrypted document from a file
    auto loadOptions = MakeObject<LoadOptions>(u"docPassword");
    ASSERT_EQ(signOptions->get_DecryptionPassword(), loadOptions->get_Password());

    // Check that encrypted document was successfully signed
    auto signedDoc = MakeObject<Document>(outputFileName, loadOptions);
    SharedPtr<DigitalSignatureCollection> signatures = signedDoc->get_DigitalSignatures();

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

void ExDigitalSignatureUtil::NoArgumentsForSing()
{
    auto signOptions = MakeObject<SignOptions>();
    signOptions->set_Comments(String::Empty);
    signOptions->set_SignTime(System::DateTime::get_Now());
    signOptions->set_DecryptionPassword(String::Empty);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_1 = [&signOptions]()
    {
        DigitalSignatureUtil::Sign(String::Empty, String::Empty, nullptr, signOptions);
    };

    ASSERT_THROW(_local_func_1(), System::ArgumentException);
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
    auto doc = MakeObject<Document>(MyDir + u"Digitally signed.docx");
    String outputFileName = ArtifactsDir + u"DigitalSignatureUtil.NoCertificateForSign.docx";

    auto signOptions = MakeObject<SignOptions>();
    signOptions->set_Comments(u"Comment");
    signOptions->set_SignTime(System::DateTime::get_Now());
    signOptions->set_DecryptionPassword(u"docPassword");

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_2 = [&doc, &outputFileName, &signOptions]()
    {
        DigitalSignatureUtil::Sign(doc->get_OriginalFileName(), outputFileName, nullptr, signOptions);
    };

    ASSERT_THROW(_local_func_2(), System::ArgumentNullException);
}

namespace gtest_test
{

TEST_F(ExDigitalSignatureUtil, NoCertificateForSign)
{
    s_instance->NoCertificateForSign();
}

} // namespace gtest_test

} // namespace ApiExamples
