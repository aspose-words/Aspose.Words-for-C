#include "Working with FileFormat.h"

using namespace Aspose::Words;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace gtest_test {

class WorkingWithFileFormat : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::WorkingWithFileFormat> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::WorkingWithFileFormat>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::WorkingWithFileFormat> WorkingWithFileFormat::s_instance;

TEST_F(WorkingWithFileFormat, DetectFileFormat)
{
    s_instance->DetectFileFormat();
}

TEST_F(WorkingWithFileFormat, DetectDocumentSignatures)
{
    s_instance->DetectDocumentSignatures();
}

TEST_F(WorkingWithFileFormat, VerifyEncryptedDocument)
{
    s_instance->VerifyEncryptedDocument();
}

}}} // namespace DocsExamples::File_Formats_and_Conversions::gtest_test
