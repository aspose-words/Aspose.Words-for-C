#include "Split document.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace DocsExamples { namespace Programming_with_Documents { namespace Split_Documents { namespace gtest_test {

class SplitDocument : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Split_Documents::SplitDocument> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Split_Documents::SplitDocument>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Split_Documents::SplitDocument> SplitDocument::s_instance;

TEST_F(SplitDocument, ByHeadingsHtml)
{
    s_instance->ByHeadingsHtml();
}

TEST_F(SplitDocument, BySectionsHtml)
{
    s_instance->BySectionsHtml();
}

TEST_F(SplitDocument, BySections)
{
    s_instance->BySections();
}

TEST_F(SplitDocument, PageByPage)
{
    s_instance->PageByPage();
}

TEST_F(SplitDocument, ByPageRange)
{
    s_instance->ByPageRange();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Split_Documents::gtest_test
