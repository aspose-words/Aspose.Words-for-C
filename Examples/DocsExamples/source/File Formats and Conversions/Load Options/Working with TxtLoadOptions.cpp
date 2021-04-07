#include "Working with TxtLoadOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Load_Options { namespace gtest_test {

class WorkingWithTxtLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Load_Options::WorkingWithTxtLoadOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::Load_Options::WorkingWithTxtLoadOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Load_Options::WorkingWithTxtLoadOptions> WorkingWithTxtLoadOptions::s_instance;

TEST_F(WorkingWithTxtLoadOptions, DetectNumberingWithWhitespaces)
{
    s_instance->DetectNumberingWithWhitespaces();
}

TEST_F(WorkingWithTxtLoadOptions, HandleSpacesOptions)
{
    s_instance->HandleSpacesOptions();
}

TEST_F(WorkingWithTxtLoadOptions, DocumentTextDirection)
{
    s_instance->DocumentTextDirection();
}

}}}} // namespace DocsExamples::File_Formats_and_Conversions::Load_Options::gtest_test
