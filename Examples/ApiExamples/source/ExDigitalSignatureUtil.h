#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/CertificateHolder.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignature.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureType.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Model/Loading/LoadOptions.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/exceptions.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/stream.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Loading;

namespace ApiExamples {

class ExDigitalSignatureUtil : public ApiExampleBase
{
public:
    void Load()
    {
        //ExStart
        //ExFor:DigitalSignatureUtil
        //ExFor:DigitalSignatureUtil.LoadSignatures(String)
        //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
        //ExSummary:Shows how to load signatures from a digitally signed document.
        // There are two ways of loading a signed document's collection of digital signatures using the DigitalSignatureUtil class.
        // 1 -  Load from a document from a local file system filename:
        SharedPtr<DigitalSignatureCollection> digitalSignatures = DigitalSignatureUtil::LoadSignatures(MyDir + u"Digitally signed.docx");

        // If this collection is nonempty, then we can verify that the document is digitally signed.
        ASSERT_EQ(1, digitalSignatures->get_Count());

        // 2 -  Load from a document from a FileStream:
        {
            SharedPtr<System::IO::Stream> stream = MakeObject<System::IO::FileStream>(MyDir + u"Digitally signed.docx", System::IO::FileMode::Open);
            digitalSignatures = DigitalSignatureUtil::LoadSignatures(stream);
            ASSERT_EQ(1, digitalSignatures->get_Count());
        }
        //ExEnd
    }

    void Remove()
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
        DigitalSignatureUtil::RemoveAllSignatures(MyDir + u"Digitally signed.docx", ArtifactsDir + u"DigitalSignatureUtil.LoadAndRemove.FromString.docx");

        // 2 - Determine the locations of both the signed document and the unsigned copy by file streams:
        {
            SharedPtr<System::IO::Stream> streamIn = MakeObject<System::IO::FileStream>(MyDir + u"Digitally signed.docx", System::IO::FileMode::Open);
            {
                SharedPtr<System::IO::Stream> streamOut =
                    MakeObject<System::IO::FileStream>(ArtifactsDir + u"DigitalSignatureUtil.LoadAndRemove.FromStream.docx", System::IO::FileMode::Create);
                DigitalSignatureUtil::RemoveAllSignatures(streamIn, streamOut);
            }
        }

        // Verify that both our output documents have no digital signatures.
        ASSERT_EQ(0, DigitalSignatureUtil::LoadSignatures(ArtifactsDir + u"DigitalSignatureUtil.LoadAndRemove.FromString.docx")->get_Count());
        ASSERT_EQ(0, DigitalSignatureUtil::LoadSignatures(ArtifactsDir + u"DigitalSignatureUtil.LoadAndRemove.FromStream.docx")->get_Count());
        //ExEnd
    }

    void SignDocument()
    {
        //ExStart
        //ExFor:CertificateHolder
        //ExFor:CertificateHolder.Create(String, String)
        //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, SignOptions)
        //ExFor:SignOptions.Comments
        //ExFor:SignOptions.SignTime
        //ExSummary:Shows how to digitally sign documents.
        // Create an X.509 certificate from a PKCS#12 store, which should contain a private key.
        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        // Create a comment and date which will be applied with our new digital signature.
        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_Comments(u"My comment");
        signOptions->set_SignTime(System::DateTime::get_Now());

        // Take an unsigned document from the local file system via a file stream,
        // then create a signed copy of it determined by the filename of the output file stream.
        {
            SharedPtr<System::IO::Stream> streamIn = MakeObject<System::IO::FileStream>(MyDir + u"Document.docx", System::IO::FileMode::Open);
            {
                SharedPtr<System::IO::Stream> streamOut =
                    MakeObject<System::IO::FileStream>(ArtifactsDir + u"DigitalSignatureUtil.SignDocument.docx", System::IO::FileMode::OpenOrCreate);
                DigitalSignatureUtil::Sign(streamIn, streamOut, certificateHolder, signOptions);
            }
        }
        //ExEnd

        {
            SharedPtr<System::IO::Stream> stream =
                MakeObject<System::IO::FileStream>(ArtifactsDir + u"DigitalSignatureUtil.SignDocument.docx", System::IO::FileMode::Open);
            SharedPtr<DigitalSignatureCollection> digitalSignatures = DigitalSignatureUtil::LoadSignatures(stream);
            ASSERT_EQ(1, digitalSignatures->get_Count());

            SharedPtr<DigitalSignature> signature = digitalSignatures->idx_get(0);

            ASSERT_TRUE(signature->get_IsValid());
            ASSERT_EQ(DigitalSignatureType::XmlDsig, signature->get_SignatureType());
            ASSERT_EQ(System::ObjectExt::ToString(signOptions->get_SignTime()), System::ObjectExt::ToString(signature->get_SignTime()));
            ASSERT_EQ(u"My comment", signature->get_Comments());
        }
    }

    void DecryptionPassword()
    {
        //ExStart
        //ExFor:CertificateHolder
        //ExFor:SignOptions.DecryptionPassword
        //ExFor:LoadOptions.Password
        //ExSummary:Shows how to sign encrypted document file.
        // Create an X.509 certificate from a PKCS#12 store, which should contain a private key.
        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        // Create a comment, date, and decryption password which will be applied with our new digital signature.
        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_Comments(u"Comment");
        signOptions->set_SignTime(System::DateTime::get_Now());
        signOptions->set_DecryptionPassword(u"docPassword");

        // Set a local system filename for the unsigned input document, and an output filename for its new digitally signed copy.
        String inputFileName = MyDir + u"Encrypted.docx";
        String outputFileName = ArtifactsDir + u"DigitalSignatureUtil.DecryptionPassword.docx";

        DigitalSignatureUtil::Sign(inputFileName, outputFileName, certificateHolder, signOptions);
        //ExEnd

        // Open encrypted document from a file.
        auto loadOptions = MakeObject<LoadOptions>(u"docPassword");
        ASSERT_EQ(signOptions->get_DecryptionPassword(), loadOptions->get_Password());

        // Check that encrypted document was successfully signed.
        auto signedDoc = MakeObject<Document>(outputFileName, loadOptions);
        SharedPtr<DigitalSignatureCollection> signatures = signedDoc->get_DigitalSignatures();

        ASSERT_EQ(1, signatures->get_Count());
        ASSERT_TRUE(signatures->get_IsValid());
    }

    void SignDocumentObfuscationBug()
    {
        SharedPtr<CertificateHolder> ch = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        auto doc = MakeObject<Document>(MyDir + u"Structured document tags.docx");
        String outputFileName = ArtifactsDir + u"DigitalSignatureUtil.SignDocumentObfuscationBug.doc";

        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_Comments(u"Comment");
        signOptions->set_SignTime(System::DateTime::get_Now());

        DigitalSignatureUtil::Sign(doc->get_OriginalFileName(), outputFileName, ch, signOptions);
    }

    void IncorrectDecryptionPassword()
    {
        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        auto doc = MakeObject<Document>(MyDir + u"Encrypted.docx", MakeObject<LoadOptions>(u"docPassword"));
        String outputFileName = ArtifactsDir + u"DigitalSignatureUtil.IncorrectDecryptionPassword.docx";

        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_Comments(u"Comment");
        signOptions->set_SignTime(System::DateTime::get_Now());
        signOptions->set_DecryptionPassword(u"docPassword1");

        ASSERT_THROW(DigitalSignatureUtil::Sign(doc->get_OriginalFileName(), outputFileName, certificateHolder, signOptions), IncorrectPasswordException);
    }

    void NoArgumentsForSing()
    {
        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_Comments(String::Empty);
        signOptions->set_SignTime(System::DateTime::get_Now());
        signOptions->set_DecryptionPassword(String::Empty);

        ASSERT_THROW(DigitalSignatureUtil::Sign(String::Empty, String::Empty, nullptr, signOptions), System::ArgumentException);
    }

    void NoCertificateForSign()
    {
        auto doc = MakeObject<Document>(MyDir + u"Digitally signed.docx");
        String outputFileName = ArtifactsDir + u"DigitalSignatureUtil.NoCertificateForSign.docx";

        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_Comments(u"Comment");
        signOptions->set_SignTime(System::DateTime::get_Now());
        signOptions->set_DecryptionPassword(u"docPassword");

        ASSERT_THROW(DigitalSignatureUtil::Sign(doc->get_OriginalFileName(), outputFileName, nullptr, signOptions), System::ArgumentNullException);
    }
};

} // namespace ApiExamples
