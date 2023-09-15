#include "Nested MailMerge custom.h"

#include "Mail Merge and Reporting/Complex examples and helpers/Nested MailMerge custom.h"

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
namespace DocsExamples { namespace Mail_Merge_and_Reporting { namespace Custom_examples { namespace gtest_test {

class NestedMailMergeCustom : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Mail_Merge_and_Reporting::Custom_examples::NestedMailMergeCustom> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Mail_Merge_and_Reporting::Custom_examples::NestedMailMergeCustom>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Mail_Merge_and_Reporting::Custom_examples::NestedMailMergeCustom> NestedMailMergeCustom::s_instance;

TEST_F(NestedMailMergeCustom, NestedMailMerge)
{
    s_instance->NestedMailMerge();
}

}}}} // namespace DocsExamples::Mail_Merge_and_Reporting::Custom_examples::gtest_test
