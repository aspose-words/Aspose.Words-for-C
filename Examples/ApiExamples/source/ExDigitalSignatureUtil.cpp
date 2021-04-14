// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDigitalSignatureUtil.h"

using namespace Aspose::Words;
using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Loading;
namespace ApiExamples { namespace gtest_test {

class ExDigitalSignatureUtil : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExDigitalSignatureUtil> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExDigitalSignatureUtil>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExDigitalSignatureUtil> ExDigitalSignatureUtil::s_instance;

TEST_F(ExDigitalSignatureUtil, Load)
{
    s_instance->Load();
}

TEST_F(ExDigitalSignatureUtil, Remove)
{
    s_instance->Remove();
}

TEST_F(ExDigitalSignatureUtil, SignDocument)
{
    s_instance->SignDocument();
}

TEST_F(ExDigitalSignatureUtil, DecryptionPassword)
{
    s_instance->DecryptionPassword();
}

TEST_F(ExDigitalSignatureUtil, SignDocumentObfuscationBug)
{
    s_instance->SignDocumentObfuscationBug();
}

TEST_F(ExDigitalSignatureUtil, IncorrectDecryptionPassword)
{
    s_instance->IncorrectDecryptionPassword();
}

TEST_F(ExDigitalSignatureUtil, NoArgumentsForSing)
{
    s_instance->NoArgumentsForSing();
}

TEST_F(ExDigitalSignatureUtil, NoCertificateForSign)
{
    s_instance->NoCertificateForSign();
}

}} // namespace ApiExamples::gtest_test
