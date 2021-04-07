#include "Working with DocSaveOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options { namespace gtest_test {

class WorkingWithDocSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithDocSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithDocSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithDocSaveOptions> WorkingWithDocSaveOptions::s_instance;

TEST_F(WorkingWithDocSaveOptions, EncryptDocumentWithPassword)
{
    s_instance->EncryptDocumentWithPassword();
}

TEST_F(WorkingWithDocSaveOptions, DoNotCompressSmallMetafiles)
{
    s_instance->DoNotCompressSmallMetafiles();
}

TEST_F(WorkingWithDocSaveOptions, DoNotSavePictureBullet)
{
    s_instance->DoNotSavePictureBullet();
}

}}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options::gtest_test
