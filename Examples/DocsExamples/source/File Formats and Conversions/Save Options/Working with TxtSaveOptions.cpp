#include "Working with TxtSaveOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options { namespace gtest_test {

class WorkingWithTxtSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithTxtSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithTxtSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithTxtSaveOptions> WorkingWithTxtSaveOptions::s_instance;

TEST_F(WorkingWithTxtSaveOptions, AddBidiMarks)
{
    s_instance->AddBidiMarks();
}

TEST_F(WorkingWithTxtSaveOptions, UseTabCharacterPerLevelForListIndentation)
{
    s_instance->UseTabCharacterPerLevelForListIndentation();
}

TEST_F(WorkingWithTxtSaveOptions, UseSpaceCharacterPerLevelForListIndentation)
{
    s_instance->UseSpaceCharacterPerLevelForListIndentation();
}

}}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options::gtest_test
