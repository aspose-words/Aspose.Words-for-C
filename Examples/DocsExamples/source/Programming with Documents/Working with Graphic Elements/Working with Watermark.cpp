#include "Working with Watermark.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Graphic_Elements { namespace gtest_test {

class WorkWithWatermark : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkWithWatermark> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkWithWatermark>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkWithWatermark> WorkWithWatermark::s_instance;

TEST_F(WorkWithWatermark, AddTextWatermarkWithSpecificOptions)
{
    s_instance->AddTextWatermarkWithSpecificOptions();
}

TEST_F(WorkWithWatermark, AddImageWatermarkWithSpecificOptions)
{
    s_instance->AddImageWatermarkWithSpecificOptions();
}

TEST_F(WorkWithWatermark, RemoveWatermarkFromDocument)
{
    s_instance->RemoveWatermarkFromDocument();
}

TEST_F(WorkWithWatermark, AddAndRemoveWatermark)
{
    s_instance->AddAndRemoveWatermark();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::gtest_test
