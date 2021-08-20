#include "Working with Markdown.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
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

TEST_F(WorkingWithMarkdown, BoldText)
{
    s_instance->BoldText();
}

TEST_F(WorkingWithMarkdown, ItalicText)
{
    s_instance->ItalicText();
}

TEST_F(WorkingWithMarkdown, Strikethrough)
{
    s_instance->Strikethrough();
}

TEST_F(WorkingWithMarkdown, InlineCode)
{
    s_instance->InlineCode();
}

TEST_F(WorkingWithMarkdown, Autolink)
{
    s_instance->Autolink();
}

TEST_F(WorkingWithMarkdown, Link)
{
    s_instance->Link();
}

TEST_F(WorkingWithMarkdown, Image)
{
    s_instance->Image();
}

TEST_F(WorkingWithMarkdown, HorizontalRule)
{
    s_instance->HorizontalRule();
}

TEST_F(WorkingWithMarkdown, Heading)
{
    s_instance->Heading();
}

TEST_F(WorkingWithMarkdown, SetextHeading)
{
    s_instance->SetextHeading();
}

TEST_F(WorkingWithMarkdown, IndentedCode)
{
    s_instance->IndentedCode();
}

TEST_F(WorkingWithMarkdown, FencedCode)
{
    s_instance->FencedCode();
}

TEST_F(WorkingWithMarkdown, Quote)
{
    s_instance->Quote();
}

TEST_F(WorkingWithMarkdown, BulletedList)
{
    s_instance->BulletedList();
}

TEST_F(WorkingWithMarkdown, OrderedList)
{
    s_instance->OrderedList();
}

TEST_F(WorkingWithMarkdown, Table)
{
    s_instance->Table();
}

TEST_F(WorkingWithMarkdown, ReadMarkdownDocument)
{
    s_instance->ReadMarkdownDocument();
}

TEST_F(WorkingWithMarkdown, Emphases)
{
    s_instance->Emphases();
}

TEST_F(WorkingWithMarkdown, UseWarningSource)
{
    s_instance->UseWarningSource();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
