// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSignDocumentCustom.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/predicate.h>
#include <system/object.h>
#include <system/io/memory_stream.h>
#include <system/guid.h>
#include <system/details/dispose_guard.h>
#include <system/console.h>
#include <system/collections/list.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/imaging/image_format.h>
#include <drawing/image.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Drawing/SignatureLine.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/SignOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SignatureLineOptions.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/Document/CertificateHolder.h>

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

SharedPtr<System::Collections::Generic::List<SharedPtr<TestData::TestClasses::SignPersonTestClass>>> ExSignDocumentCustom::gSignPersonList;

void ExSignDocumentCustom::SignDocument(String srcDocumentPath, String dstDocumentPath, SharedPtr<TestData::TestClasses::SignPersonTestClass> signPersonInfo, String certificatePath, String certificatePassword)
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

ArrayPtr<uint8_t> ExSignDocumentCustom::ImageToByteArray(SharedPtr<System::Drawing::Image> imageIn)
{
    {
        auto ms = MakeObject<System::IO::MemoryStream>();
        imageIn->Save(ms, System::Drawing::Imaging::ImageFormat::get_Png());
        return ms->ToArray();
    }
}

void ExSignDocumentCustom::CreateSignPersonData()
{
    gSignPersonList = [&]{ SharedPtr<ApiExamples::TestData::TestClasses::SignPersonTestClass> init_0[] = {MakeObject<TestData::TestClasses::SignPersonTestClass>(System::Guid::NewGuid(), u"Ron Williams", u"Chief Executive Officer", ImageToByteArray(System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg"))), MakeObject<TestData::TestClasses::SignPersonTestClass>(System::Guid::NewGuid(), u"Stephen Morse", u"Head of Compliance", ImageToByteArray(System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg")))}; auto list_0 = MakeObject<System::Collections::Generic::List<SharedPtr<TestData::TestClasses::SignPersonTestClass>>>(); list_0->AddInitializer(2, init_0); return list_0; }();
}

namespace gtest_test
{

class ExSignDocumentCustom : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExSignDocumentCustom> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExSignDocumentCustom>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExSignDocumentCustom> ExSignDocumentCustom::s_instance;

} // namespace gtest_test

void ExSignDocumentCustom::Sign()
{
    String signPersonName = u"Ron Williams";
    String srcDocumentPath = MyDir + u"Document.docx";
    String dstDocumentPath = ArtifactsDir + u"SignDocumentCustom.Sign.docx";
    String certificatePath = MyDir + u"morzal.pfx";
    String certificatePassword = u"aw";

    // We need to create simple list with test signers for this example
    CreateSignPersonData();
    System::Console::WriteLine(u"Test data successfully added!");

    // Get sign person object by name of the person who must sign a document
    // This an example, in real use case you would return an object from a database

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<TestData::TestClasses::SignPersonTestClass> c)> _local_func_0 = [&signPersonName](SharedPtr<TestData::TestClasses::SignPersonTestClass> c)
    {
        return c->get_Name() == signPersonName;
    };

    SharedPtr<TestData::TestClasses::SignPersonTestClass> signPersonInfo = gSignPersonList->Find(static_cast<System::Predicate<SharedPtr<TestData::TestClasses::SignPersonTestClass>>>(_local_func_0));

    if (signPersonInfo != nullptr)
    {
        SignDocument(srcDocumentPath, dstDocumentPath, signPersonInfo, certificatePath, certificatePassword);
        System::Console::WriteLine(u"Document successfully signed!");
    }
    else
    {
        System::Console::WriteLine(u"Sign person does not exist, please check your parameters.");
        FAIL();
        //ExSkip
    }

    // Now do something with a signed document, for example, save it to your database
    // Use 'new Document(dstDocumentPath)' for loading a signed document
}

namespace gtest_test
{

TEST_F(ExSignDocumentCustom, Sign)
{
    s_instance->Sign();
}

} // namespace gtest_test

} // namespace ApiExamples
