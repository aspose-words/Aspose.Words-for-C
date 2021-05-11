#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/DigitalSignatures/CertificateHolder.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/SignatureLine.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/SignatureLineOptions.h>
#include <drawing/image.h>
#include <drawing/imaging/image_format.h>
#include <system/array.h>
#include <system/collections/list.h>
#include <system/details/dispose_guard.h>
#include <system/guid.h>
#include <system/io/memory_stream.h>
#include <system/predicate.h>
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
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::DigitalSignatures;

namespace ApiExamples {

class ExSignDocumentCustom : public ApiExampleBase
{
public:
    //ExStart
    //ExFor:CertificateHolder
    //ExFor:SignatureLineOptions.Signer
    //ExFor:SignatureLineOptions.SignerTitle
    //ExFor:SignatureLine.Id
    //ExFor:SignOptions.SignatureLineId
    //ExFor:SignOptions.SignatureLineImage
    //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
    //ExSummary:Shows how to add a signature line to a document, and then sign it using a digital certificate.
    static void Sign()
    {
        String signeeName = u"Ron Williams";
        String srcDocumentPath = MyDir + u"Document.docx";
        String dstDocumentPath = ArtifactsDir + u"SignDocumentCustom.Sign.docx";
        String certificatePath = MyDir + u"morzal.pfx";
        String certificatePassword = u"aw";

        CreateSignees();

        auto signeeInfo = mSignees->Find(System::Predicate<SharedPtr<Signee>>([&](SharedPtr<Signee> c) { return c->get_Name() == signeeName; }));

        if (signeeInfo != nullptr)
        {
            SignDocument(srcDocumentPath, dstDocumentPath, signeeInfo, certificatePath, certificatePassword);
        }
        else
        {
            FAIL() << "Signee does not exist.";
        }
    }

    class Signee : public System::Object
    {
    public:
        System::Guid get_PersonId()
        {
            return pr_PersonId;
        }

        void set_PersonId(System::Guid value)
        {
            pr_PersonId = value;
        }

        String get_Name()
        {
            return pr_Name;
        }

        void set_Name(String value)
        {
            pr_Name = value;
        }

        String get_Position()
        {
            return pr_Position;
        }

        void set_Position(String value)
        {
            pr_Position = value;
        }

        ArrayPtr<uint8_t> get_Image()
        {
            return pr_Image;
        }

        void set_Image(ArrayPtr<uint8_t> value)
        {
            pr_Image = value;
        }

        Signee(System::Guid guid, String name, String position, ArrayPtr<uint8_t> image)
        {
            set_PersonId(guid);
            set_Name(name);
            set_Position(position);
            set_Image(image);
        }

    private:
        System::Guid pr_PersonId;
        String pr_Name;
        String pr_Position;
        ArrayPtr<uint8_t> pr_Image;
    };
    /// <summary>
    /// Creates a copy of a source document signed using provided signee information and X509 certificate.
    /// </summary>
    static void SignDocument(String srcDocumentPath, String dstDocumentPath, SharedPtr<ExSignDocumentCustom::Signee> signeeInfo, String certificatePath,
                             String certificatePassword)
    {
        auto document = MakeObject<Document>(srcDocumentPath);
        auto builder = MakeObject<DocumentBuilder>(document);

        // Configure and insert a signature line, an object in the document that will display a signature that we sign it with.
        auto signatureLineOptions = MakeObject<SignatureLineOptions>();
        signatureLineOptions->set_Signer(signeeInfo->get_Name());
        signatureLineOptions->set_SignerTitle(signeeInfo->get_Position());

        SharedPtr<SignatureLine> signatureLine = builder->InsertSignatureLine(signatureLineOptions)->get_SignatureLine();
        signatureLine->set_Id(signeeInfo->get_PersonId());

        // First, we will save an unsigned version of our document.
        builder->get_Document()->Save(dstDocumentPath);

        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(certificatePath, certificatePassword);

        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_SignatureLineId(signeeInfo->get_PersonId());
        signOptions->set_SignatureLineImage(signeeInfo->get_Image());

        // Overwrite the unsigned document we saved above with a version signed using the certificate.
        DigitalSignatureUtil::Sign(dstDocumentPath, dstDocumentPath, certificateHolder, signOptions);
    }

    /// <summary>
    /// Converts an image to a byte array.
    /// </summary>
    static ArrayPtr<uint8_t> ImageToByteArray(SharedPtr<System::Drawing::Image> imageIn)
    {
        {
            auto ms = MakeObject<System::IO::MemoryStream>();
            imageIn->Save(ms, System::Drawing::Imaging::ImageFormat::get_Png());
            return ms->ToArray();
        }
    }

    static void CreateSignees()
    {
        mSignees = MakeObject<System::Collections::Generic::List<SharedPtr<ExSignDocumentCustom::Signee>>>();
        mSignees->Add(MakeObject<ExSignDocumentCustom::Signee>(System::Guid::NewGuid(), u"Ron Williams", u"Chief Executive Officer",
                                                               ImageToByteArray(System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg"))));
        mSignees->Add(MakeObject<ExSignDocumentCustom::Signee>(System::Guid::NewGuid(), u"Stephen Morse", u"Head of Compliance",
                                                               ImageToByteArray(System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg"))));
    }

    static SharedPtr<System::Collections::Generic::List<SharedPtr<ExSignDocumentCustom::Signee>>> mSignees;
    //ExEnd
};

} // namespace ApiExamples
