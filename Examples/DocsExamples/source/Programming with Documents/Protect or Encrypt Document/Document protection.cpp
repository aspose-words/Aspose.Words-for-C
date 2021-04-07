#include "Document protection.h"

using namespace Aspose::Words;
namespace DocsExamples { namespace Programming_with_Documents { namespace Protect_or_Encrypt_Document { namespace gtest_test {

class DocumentProtection : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Protect_or_Encrypt_Document::DocumentProtection> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Protect_or_Encrypt_Document::DocumentProtection>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Protect_or_Encrypt_Document::DocumentProtection> DocumentProtection::s_instance;

TEST_F(DocumentProtection, Protect)
{
    s_instance->Protect();
}

TEST_F(DocumentProtection, Unprotect)
{
    s_instance->Unprotect();
}

TEST_F(DocumentProtection, GetProtectionType)
{
    s_instance->GetProtectionType();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Protect_or_Encrypt_Document::gtest_test
