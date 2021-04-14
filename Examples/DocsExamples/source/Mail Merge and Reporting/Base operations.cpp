// CPPDEFECT: System.Data is not supported
#include "Base operations.h"

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
namespace DocsExamples { namespace Mail_Merge_and_Reporting { namespace gtest_test {

class BaseOperations : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Mail_Merge_and_Reporting::BaseOperations> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Mail_Merge_and_Reporting::BaseOperations>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Mail_Merge_and_Reporting::BaseOperations> BaseOperations::s_instance;

TEST_F(BaseOperations, SimpleMailMerge)
{
    s_instance->SimpleMailMerge();
}

TEST_F(BaseOperations, UseIfElseMustache)
{
    s_instance->UseIfElseMustache();
}

TEST_F(BaseOperations, GetRegionsByName)
{
    s_instance->GetRegionsByName();
}

}}} // namespace DocsExamples::Mail_Merge_and_Reporting::gtest_test
