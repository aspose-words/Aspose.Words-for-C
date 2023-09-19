#include "Working with DigitalSinatures.h"

using namespace Aspose::Words;
using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;
namespace DocsExamples { namespace Programming_with_Documents { namespace Protect_or_Encrypt_Document { namespace gtest_test {

class WorkingWithDigitalSinatures : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Protect_or_Encrypt_Document::WorkingWithDigitalSinatures> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Protect_or_Encrypt_Document::WorkingWithDigitalSinatures>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Protect_or_Encrypt_Document::WorkingWithDigitalSinatures> WorkingWithDigitalSinatures::s_instance;

TEST_F(WorkingWithDigitalSinatures, SignDocument)
{
    s_instance->SignDocument();
}

TEST_F(WorkingWithDigitalSinatures, SigningEncryptedDocument)
{
    s_instance->SigningEncryptedDocument();
}

TEST_F(WorkingWithDigitalSinatures, CreatingAndSigningNewSignatureLine)
{
    s_instance->CreatingAndSigningNewSignatureLine();
}

TEST_F(WorkingWithDigitalSinatures, SigningExistingSignatureLine)
{
    s_instance->SigningExistingSignatureLine();
}

TEST_F(WorkingWithDigitalSinatures, SetSignatureProviderId)
{
    s_instance->SetSignatureProviderId();
}

TEST_F(WorkingWithDigitalSinatures, CreateNewSignatureLineAndSetProviderId)
{
    s_instance->CreateNewSignatureLineAndSetProviderId();
}

TEST_F(WorkingWithDigitalSinatures, RemoveSignatures)
{
    s_instance->RemoveSignatures();
}

TEST_F(WorkingWithDigitalSinatures, SignatureValue)
{
    s_instance->SignatureValue();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Protect_or_Encrypt_Document::gtest_test
