#include "Working with RtfLoadOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Load_Options { namespace gtest_test {

class WorkingWithRtfLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Load_Options::WorkingWithRtfLoadOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::Load_Options::WorkingWithRtfLoadOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Load_Options::WorkingWithRtfLoadOptions> WorkingWithRtfLoadOptions::s_instance;

TEST_F(WorkingWithRtfLoadOptions, RecognizeUtf8Text)
{
    s_instance->RecognizeUtf8Text();
}

}}}} // namespace DocsExamples::File_Formats_and_Conversions::Load_Options::gtest_test
