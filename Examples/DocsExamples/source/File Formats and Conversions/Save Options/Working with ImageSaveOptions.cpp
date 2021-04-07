#include "Working with ImageSaveOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options { namespace gtest_test {

class WorkingWithImageSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithImageSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithImageSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithImageSaveOptions> WorkingWithImageSaveOptions::s_instance;

TEST_F(WorkingWithImageSaveOptions, ExposeThresholdControlForTiffBinarization)
{
    s_instance->ExposeThresholdControlForTiffBinarization();
}

TEST_F(WorkingWithImageSaveOptions, GetTiffPageRange)
{
    s_instance->GetTiffPageRange();
}

TEST_F(WorkingWithImageSaveOptions, Format1BppIndexed)
{
    s_instance->Format1BppIndexed();
}

TEST_F(WorkingWithImageSaveOptions, GetJpegPageRange)
{
    s_instance->GetJpegPageRange();
}

TEST_F(WorkingWithImageSaveOptions, PageSavingCallback)
{
    s_instance->PageSavingCallback();
}

}}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options::gtest_test
