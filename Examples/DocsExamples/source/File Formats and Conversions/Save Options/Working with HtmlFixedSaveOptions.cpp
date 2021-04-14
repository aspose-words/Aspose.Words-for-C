#include "Working with HtmlFixedSaveOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options { namespace gtest_test {

class WorkingWithHtmlFixedSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithHtmlFixedSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithHtmlFixedSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithHtmlFixedSaveOptions> WorkingWithHtmlFixedSaveOptions::s_instance;

TEST_F(WorkingWithHtmlFixedSaveOptions, UseFontFromTargetMachine)
{
    s_instance->UseFontFromTargetMachine();
}

TEST_F(WorkingWithHtmlFixedSaveOptions, WriteAllCssRulesInSingleFile)
{
    s_instance->WriteAllCssRulesInSingleFile();
}

}}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options::gtest_test
