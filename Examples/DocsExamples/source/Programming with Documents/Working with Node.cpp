#include "Working with Node.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithNode : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithNode> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithNode>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithNode> WorkingWithNode::s_instance;

TEST_F(WorkingWithNode, UseNodeType)
{
    s_instance->UseNodeType();
}

TEST_F(WorkingWithNode, GetParentNode)
{
    s_instance->GetParentNode();
}

TEST_F(WorkingWithNode, OwnerDocument)
{
    s_instance->OwnerDocument();
}

TEST_F(WorkingWithNode, EnumerateChildNodes)
{
    s_instance->EnumerateChildNodes();
}

TEST_F(WorkingWithNode, RecurseAllNodes)
{
    s_instance->RecurseAllNodes();
}

TEST_F(WorkingWithNode, TypedAccess)
{
    s_instance->TypedAccess();
}

TEST_F(WorkingWithNode, CreateAndAddParagraphNode)
{
    s_instance->CreateAndAddParagraphNode();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
