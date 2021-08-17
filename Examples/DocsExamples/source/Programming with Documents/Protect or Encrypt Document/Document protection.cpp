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

TEST_F(DocumentProtection, PasswordProtection)
{
    s_instance->PasswordProtection();
}

TEST_F(DocumentProtection, AllowOnlyFormFieldsProtect)
{
    s_instance->AllowOnlyFormFieldsProtect();
}

TEST_F(DocumentProtection, RemoveDocumentProtection)
{
    s_instance->RemoveDocumentProtection();
}

TEST_F(DocumentProtection, UnrestrictedEditableRegions)
{
    s_instance->UnrestrictedEditableRegions();
}

TEST_F(DocumentProtection, UnrestrictedSection)
{
    s_instance->UnrestrictedSection();
}

TEST_F(DocumentProtection, GetProtectionType)
{
    s_instance->GetProtectionType();
}

TEST_F(DocumentProtection, ReadOnlyProtection)
{
    s_instance->ReadOnlyProtection();
}

TEST_F(DocumentProtection, RemoveReadOnlyRestriction)
{
    s_instance->RemoveReadOnlyRestriction();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Protect_or_Encrypt_Document::gtest_test
