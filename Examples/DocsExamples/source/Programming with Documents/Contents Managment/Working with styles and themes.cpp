#include "Working with styles and themes.h"

using namespace Aspose::Words;
namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Managment { namespace gtest_test {

class WorkingWithStylesAndThemes : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Managment::WorkingWithStylesAndThemes> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Contents_Managment::WorkingWithStylesAndThemes>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Managment::WorkingWithStylesAndThemes> WorkingWithStylesAndThemes::s_instance;

TEST_F(WorkingWithStylesAndThemes, AccessStyles)
{
    s_instance->AccessStyles();
}

TEST_F(WorkingWithStylesAndThemes, CopyStyles)
{
    s_instance->CopyStyles();
}

TEST_F(WorkingWithStylesAndThemes, GetThemeProperties)
{
    s_instance->GetThemeProperties();
}

TEST_F(WorkingWithStylesAndThemes, SetThemeProperties)
{
    s_instance->SetThemeProperties();
}

TEST_F(WorkingWithStylesAndThemes, InsertStyleSeparator)
{
    s_instance->InsertStyleSeparator();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Contents_Managment::gtest_test
