#include "Working with MarkdownSaveOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options { namespace gtest_test {

class WorkingWithMarkdownSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithMarkdownSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithMarkdownSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithMarkdownSaveOptions> WorkingWithMarkdownSaveOptions::s_instance;

TEST_F(WorkingWithMarkdownSaveOptions, ExportIntoMarkdownWithTableContentAlignment)
{
    s_instance->ExportIntoMarkdownWithTableContentAlignment();
}

TEST_F(WorkingWithMarkdownSaveOptions, SetImagesFolder)
{
    s_instance->SetImagesFolder();
}

}}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options::gtest_test
