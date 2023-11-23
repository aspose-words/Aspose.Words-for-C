#include "Working with Headers and Footers.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Tables;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithHeadersAndFooters : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithHeadersAndFooters> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithHeadersAndFooters>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithHeadersAndFooters> WorkingWithHeadersAndFooters::s_instance;

TEST_F(WorkingWithHeadersAndFooters, CreateHeaderFooter)
{
    s_instance->CreateHeaderFooter();
}

TEST_F(WorkingWithHeadersAndFooters, DifferentFirstPage)
{
    s_instance->DifferentFirstPage();
}

TEST_F(WorkingWithHeadersAndFooters, OddEvenPages)
{
    s_instance->OddEvenPages();
}

TEST_F(WorkingWithHeadersAndFooters, InsertImage)
{
    s_instance->InsertImage();
}

TEST_F(WorkingWithHeadersAndFooters, FontProps)
{
    s_instance->FontProps();
}

TEST_F(WorkingWithHeadersAndFooters, PageNumbers)
{
    s_instance->PageNumbers();
}

TEST_F(WorkingWithHeadersAndFooters, LinkToPreviousHeaderFooter)
{
    s_instance->LinkToPreviousHeaderFooter();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
