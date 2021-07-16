#include "Working with table of content.h"

using namespace Aspose::Words;
namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Management { namespace gtest_test {

class WorkingWithTableOfContent : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Management::WorkingWithTableOfContent> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Contents_Management::WorkingWithTableOfContent>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Management::WorkingWithTableOfContent> WorkingWithTableOfContent::s_instance;

TEST_F(WorkingWithTableOfContent, ChangeStyleOfTocLevel)
{
    s_instance->ChangeStyleOfTocLevel();
}

TEST_F(WorkingWithTableOfContent, ChangeTocTabStops)
{
    s_instance->ChangeTocTabStops();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Contents_Management::gtest_test
