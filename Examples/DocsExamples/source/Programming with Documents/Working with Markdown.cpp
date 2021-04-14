#include "Working with Markdown.h"

using namespace Aspose::Words;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithMarkdown : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithMarkdown> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithMarkdown>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithMarkdown> WorkingWithMarkdown::s_instance;

TEST_F(WorkingWithMarkdown, CreateMarkdownDocument)
{
    s_instance->CreateMarkdownDocument();
}

TEST_F(WorkingWithMarkdown, ReadMarkdownDocument)
{
    s_instance->ReadMarkdownDocument();
}

TEST_F(WorkingWithMarkdown, Emphases)
{
    s_instance->Emphases();
}

TEST_F(WorkingWithMarkdown, Headings)
{
    s_instance->Headings();
}

TEST_F(WorkingWithMarkdown, BlockQuotes)
{
    s_instance->BlockQuotes();
}

TEST_F(WorkingWithMarkdown, HorizontalRule)
{
    s_instance->HorizontalRule();
}

TEST_F(WorkingWithMarkdown, UseWarningSource)
{
    s_instance->UseWarningSource();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
