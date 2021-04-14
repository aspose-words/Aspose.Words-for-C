// CPPDEFECT: System.Data is not supported
#include "Working with Fields.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;
namespace DocsExamples { namespace Mail_Merge_and_Reporting { namespace gtest_test {

class WorkingWithFields_ : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Mail_Merge_and_Reporting::WorkingWithFields_> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Mail_Merge_and_Reporting::WorkingWithFields_>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Mail_Merge_and_Reporting::WorkingWithFields_> WorkingWithFields_::s_instance;

TEST_F(WorkingWithFields_, MailMergeFormFields)
{
    s_instance->MailMergeFormFields();
}

TEST_F(WorkingWithFields_, MailMergeImageField)
{
    s_instance->MailMergeImageField();
}

TEST_F(WorkingWithFields_, HandleMailMergeSwitches)
{
    s_instance->HandleMailMergeSwitches();
}

}}} // namespace DocsExamples::Mail_Merge_and_Reporting::gtest_test
