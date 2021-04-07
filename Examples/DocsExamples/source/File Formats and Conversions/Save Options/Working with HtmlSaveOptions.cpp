#include "Working with HtmlSaveOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options { namespace gtest_test {

class WorkingWithHtmlSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithHtmlSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithHtmlSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithHtmlSaveOptions> WorkingWithHtmlSaveOptions::s_instance;

TEST_F(WorkingWithHtmlSaveOptions, ExportRoundtripInformation)
{
    s_instance->ExportRoundtripInformation();
}

TEST_F(WorkingWithHtmlSaveOptions, ExportFontsAsBase64)
{
    s_instance->ExportFontsAsBase64();
}

TEST_F(WorkingWithHtmlSaveOptions, ExportResources)
{
    s_instance->ExportResources();
}

TEST_F(WorkingWithHtmlSaveOptions, ConvertMetafilesToEmfOrWmf)
{
    s_instance->ConvertMetafilesToEmfOrWmf();
}

TEST_F(WorkingWithHtmlSaveOptions, ConvertMetafilesToSvg)
{
    s_instance->ConvertMetafilesToSvg();
}

TEST_F(WorkingWithHtmlSaveOptions, AddCssClassNamePrefix)
{
    s_instance->AddCssClassNamePrefix();
}

TEST_F(WorkingWithHtmlSaveOptions, ExportCidUrlsForMhtmlResources)
{
    s_instance->ExportCidUrlsForMhtmlResources();
}

TEST_F(WorkingWithHtmlSaveOptions, ResolveFontNames)
{
    s_instance->ResolveFontNames();
}

TEST_F(WorkingWithHtmlSaveOptions, ExportTextInputFormFieldAsText)
{
    s_instance->ExportTextInputFormFieldAsText();
}

}}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options::gtest_test
