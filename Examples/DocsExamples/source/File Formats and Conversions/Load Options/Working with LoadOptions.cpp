#include "Working with LoadOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Load_Options { namespace gtest_test {

class WorkingWithLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Load_Options::WorkingWithLoadOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::Load_Options::WorkingWithLoadOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Load_Options::WorkingWithLoadOptions> WorkingWithLoadOptions::s_instance;

TEST_F(WorkingWithLoadOptions, UpdateDirtyFields)
{
    s_instance->UpdateDirtyFields();
}

TEST_F(WorkingWithLoadOptions, LoadEncryptedDocument)
{
    s_instance->LoadEncryptedDocument();
}

TEST_F(WorkingWithLoadOptions, ConvertShapeToOfficeMath)
{
    s_instance->ConvertShapeToOfficeMath();
}

TEST_F(WorkingWithLoadOptions, SetMsWordVersion)
{
    s_instance->SetMsWordVersion();
}

TEST_F(WorkingWithLoadOptions, UseTempFolder)
{
    s_instance->UseTempFolder();
}

TEST_F(WorkingWithLoadOptions, WarningCallback)
{
    s_instance->WarningCallback();
}

TEST_F(WorkingWithLoadOptions, ResourceLoadingCallback)
{
    s_instance->ResourceLoadingCallback();
}

TEST_F(WorkingWithLoadOptions, LoadWithEncoding)
{
    s_instance->LoadWithEncoding();
}

TEST_F(WorkingWithLoadOptions, ConvertMetafilesToPng)
{
    s_instance->ConvertMetafilesToPng();
}

TEST_F(WorkingWithLoadOptions, LoadChm)
{
    s_instance->LoadChm();
}

}}}} // namespace DocsExamples::File_Formats_and_Conversions::Load_Options::gtest_test
