#include "Working with OoxmlSaveOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options { namespace gtest_test {

class WorkingWithOoxmlSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithOoxmlSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithOoxmlSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithOoxmlSaveOptions> WorkingWithOoxmlSaveOptions::s_instance;

TEST_F(WorkingWithOoxmlSaveOptions, EncryptDocxWithPassword)
{
    s_instance->EncryptDocxWithPassword();
}

TEST_F(WorkingWithOoxmlSaveOptions, OoxmlComplianceIso29500_2008_Strict)
{
    s_instance->OoxmlComplianceIso29500_2008_Strict();
}

TEST_F(WorkingWithOoxmlSaveOptions, UpdateLastSavedTimeProperty)
{
    s_instance->UpdateLastSavedTimeProperty();
}

TEST_F(WorkingWithOoxmlSaveOptions, KeepLegacyControlChars)
{
    s_instance->KeepLegacyControlChars();
}

TEST_F(WorkingWithOoxmlSaveOptions, SetCompressionLevel)
{
    s_instance->SetCompressionLevel();
}

}}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options::gtest_test
