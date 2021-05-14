#include "Working with CleanupOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;
namespace DocsExamples { namespace Mail_Merge_and_Reporting { namespace gtest_test {

class WorkingWithCleanupOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Mail_Merge_and_Reporting::WorkingWithCleanupOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Mail_Merge_and_Reporting::WorkingWithCleanupOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Mail_Merge_and_Reporting::WorkingWithCleanupOptions> WorkingWithCleanupOptions::s_instance;

TEST_F(WorkingWithCleanupOptions, CleanupParagraphsWithPunctuationMarks)
{
    s_instance->CleanupParagraphsWithPunctuationMarks();
}

TEST_F(WorkingWithCleanupOptions, RemoveEmptyParagraphs)
{
    s_instance->RemoveEmptyParagraphs();
}

TEST_F(WorkingWithCleanupOptions, RemoveUnusedFields)
{
    s_instance->RemoveUnusedFields();
}

TEST_F(WorkingWithCleanupOptions, RemoveContainingFields)
{
    s_instance->RemoveContainingFields();
}

TEST_F(WorkingWithCleanupOptions, RemoveEmptyTableRows)
{
    s_instance->RemoveEmptyTableRows();
}

}}} // namespace DocsExamples::Mail_Merge_and_Reporting::gtest_test
