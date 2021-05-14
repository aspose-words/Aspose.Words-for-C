#include "Base conversions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace gtest_test {

class BaseConversions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::BaseConversions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::BaseConversions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::BaseConversions> BaseConversions::s_instance;

TEST_F(BaseConversions, DocToDocx)
{
    s_instance->DocToDocx();
}

TEST_F(BaseConversions, DocxToRtf)
{
    s_instance->DocxToRtf();
}

TEST_F(BaseConversions, DocxToPdf)
{
    s_instance->DocxToPdf();
}

TEST_F(BaseConversions, DocxToByte)
{
    s_instance->DocxToByte();
}

TEST_F(BaseConversions, DocxToEpub)
{
    s_instance->DocxToEpub();
}

TEST_F(BaseConversions, DocxToMarkdown)
{
    s_instance->DocxToMarkdown();
}

TEST_F(BaseConversions, DocxToTxt)
{
    s_instance->DocxToTxt();
}

TEST_F(BaseConversions, TxtToDocx)
{
    s_instance->TxtToDocx();
}

TEST_F(BaseConversions, ImagesToPdf)
{
    s_instance->ImagesToPdf();
}

}}} // namespace DocsExamples::File_Formats_and_Conversions::gtest_test
