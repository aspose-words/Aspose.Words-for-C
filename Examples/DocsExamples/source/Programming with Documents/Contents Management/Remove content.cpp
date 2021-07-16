#include "Remove content.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Management { namespace gtest_test {

class RemoveContent : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Management::RemoveContent> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Contents_Management::RemoveContent>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Management::RemoveContent> RemoveContent::s_instance;

TEST_F(RemoveContent, RemovePageBreaks)
{
    s_instance->RemovePageBreaks();
}

TEST_F(RemoveContent, RemoveFooters)
{
    s_instance->RemoveFooters();
}

TEST_F(RemoveContent, RemoveToc)
{
    s_instance->RemoveToc();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Contents_Management::gtest_test
