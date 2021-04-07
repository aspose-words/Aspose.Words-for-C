#include "Working with Comments.h"

using namespace Aspose::Words;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithComments : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithComments> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithComments>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithComments> WorkingWithComments::s_instance;

TEST_F(WorkingWithComments, AddComments)
{
    s_instance->AddComments();
}

TEST_F(WorkingWithComments, AnchorComment)
{
    s_instance->AnchorComment();
}

TEST_F(WorkingWithComments, AddRemoveCommentReply)
{
    s_instance->AddRemoveCommentReply();
}

TEST_F(WorkingWithComments, ProcessComments)
{
    s_instance->ProcessComments();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
