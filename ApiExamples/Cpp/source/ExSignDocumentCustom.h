#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/CertificateHolder.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/SignOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SignatureLineOptions.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/SignatureLine.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
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
#include "TestData/TestClasses/SignPersonTestClass.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace ApiExamples::TestData::TestClasses;
using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace ApiExamples {

/// <summary>
/// This example demonstrates how to add new signature line to the document and sign it with your personal signature <see cref="SignDocument"></see>.
/// </summary>
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
    //ExSummary:Demonstrates how to add new signature line to the document and sign it with personal signature using SignatureLineId.
    static void Sign()
    {
        String signPersonName = u"Ron Williams";
        String srcDocumentPath = MyDir + u"Document.docx";
        String dstDocumentPath = ArtifactsDir + u"SignDocumentCustom.Sign.docx";
        String certificatePath = MyDir + u"morzal.pfx";
        String certificatePassword = u"aw";

        // We need to create simple list with test signers for this example
        CreateSignPersonData();
        std::cout << "Test data successfully added!" << std::endl;

        // Get sign person object by name of the person who must sign a document
        // This an example, in real use case you would return an object from a database

        std::function<bool(SharedPtr<TestData::TestClasses::SignPersonTestClass> c)> hasSignPersonName = [&signPersonName](SharedPtr<TestData::TestClasses::SignPersonTestClass> c)
        {
            return c->get_Name() == signPersonName;
        };

        SharedPtr<TestData::TestClasses::SignPersonTestClass> signPersonInfo = gSignPersonList->Find(hasSignPersonName);

        if (signPersonInfo != nullptr)
        {
            SignDocument(srcDocumentPath, dstDocumentPath, signPersonInfo, certificatePath, certificatePassword);
            std::cout << "Document successfully signed!" << std::endl;
        }
        else
        {
            std::cout << "Sign person does not exist, please check your parameters." << std::endl;
            FAIL();
            //ExSkip
        }

        // Now do something with a signed document, for example, save it to your database
        // Use 'new Document(dstDocumentPath)' for loading a signed document
    }

protected:
    /// <summary>
    /// Signs the document obtained at the source location and saves it to the specified destination.
    /// </summary>
    static void SignDocument(String srcDocumentPath, String dstDocumentPath,
                             SharedPtr<TestData::TestClasses::SignPersonTestClass> signPersonInfo, String certificatePath,
                             String certificatePassword)
    {
        // Create new document instance based on a test file that we need to sign
        auto document = MakeObject<Document>(srcDocumentPath);
        auto builder = MakeObject<DocumentBuilder>(document);

        // Add info about responsible person who sign a document
        auto signatureLineOptions = MakeObject<SignatureLineOptions>();
        signatureLineOptions->set_Signer(signPersonInfo->get_Name());
        signatureLineOptions->set_SignerTitle(signPersonInfo->get_Position());

        // Add signature line for responsible person who sign a document
        SharedPtr<SignatureLine> signatureLine = builder->InsertSignatureLine(signatureLineOptions)->get_SignatureLine();
        signatureLine->set_Id(signPersonInfo->get_PersonId());

        // Save a document with line signatures into temporary file for future signing
        builder->get_Document()->Save(dstDocumentPath);

        // Create holder of certificate instance based on your personal certificate
        // This is the test certificate generated for this example
        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(certificatePath, certificatePassword);

        // Link our signature line with personal signature
        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_SignatureLineId(signPersonInfo->get_PersonId());
        signOptions->set_SignatureLineImage(signPersonInfo->get_Image());

        // Sign a document which contains signature line with personal certificate
        DigitalSignatureUtil::Sign(dstDocumentPath, dstDocumentPath, certificateHolder, signOptions);
    }

    /// <summary>
    /// Converting image file to bytes array.
    /// </summary>
    static ArrayPtr<uint8_t> ImageToByteArray(SharedPtr<System::Drawing::Image> imageIn)
    {
        {
            auto ms = MakeObject<System::IO::MemoryStream>();
            imageIn->Save(ms, System::Drawing::Imaging::ImageFormat::get_Png());
            return ms->ToArray();
        }
    }

    /// <summary>
    /// Create test data that contains info about sing persons
    /// </summary>
    static void CreateSignPersonData()
    {
        gSignPersonList = [&]
        {
            SharedPtr<ApiExamples::TestData::TestClasses::SignPersonTestClass> init_0[] = {
                MakeObject<TestData::TestClasses::SignPersonTestClass>(System::Guid::NewGuid(), u"Ron Williams", u"Chief Executive Officer",
                                                                               ImageToByteArray(System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg"))),
                MakeObject<TestData::TestClasses::SignPersonTestClass>(System::Guid::NewGuid(), u"Stephen Morse", u"Head of Compliance",
                                                                               ImageToByteArray(System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg")))};
            auto list_0 = MakeObject<System::Collections::Generic::List<SharedPtr<TestData::TestClasses::SignPersonTestClass>>>();
            list_0->AddInitializer(2, init_0);
            return list_0;
        }();
    }

    static SharedPtr<System::Collections::Generic::List<SharedPtr<TestData::TestClasses::SignPersonTestClass>>> gSignPersonList;
    //ExEnd
};

} // namespace ApiExamples
