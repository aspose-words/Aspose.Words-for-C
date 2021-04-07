#include "Working with Section.h"

using namespace Aspose::Words;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithSection : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithSection> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithSection>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithSection> WorkingWithSection::s_instance;

TEST_F(WorkingWithSection, AddSection)
{
    s_instance->AddSection();
}

TEST_F(WorkingWithSection, DeleteSection)
{
    s_instance->DeleteSection();
}

TEST_F(WorkingWithSection, DeleteAllSections)
{
    s_instance->DeleteAllSections();
}

TEST_F(WorkingWithSection, AppendSectionContent)
{
    s_instance->AppendSectionContent();
}

TEST_F(WorkingWithSection, CloneSection)
{
    s_instance->CloneSection();
}

TEST_F(WorkingWithSection, CopySection)
{
    s_instance->CopySection();
}

TEST_F(WorkingWithSection, DeleteHeaderFooterContent)
{
    s_instance->DeleteHeaderFooterContent();
}

TEST_F(WorkingWithSection, DeleteSectionContent)
{
    s_instance->DeleteSectionContent();
}

TEST_F(WorkingWithSection, ModifyPageSetupInAllSections)
{
    s_instance->ModifyPageSetupInAllSections();
}

TEST_F(WorkingWithSection, SectionsAccessByIndex)
{
    s_instance->SectionsAccessByIndex();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
