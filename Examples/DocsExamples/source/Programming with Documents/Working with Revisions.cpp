#include "Working with Revisions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Layout;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithRevisions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithRevisions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithRevisions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithRevisions> WorkingWithRevisions::s_instance;

TEST_F(WorkingWithRevisions, AcceptRevisions)
{
    s_instance->AcceptRevisions();
}

TEST_F(WorkingWithRevisions, GetRevisionTypes)
{
    s_instance->GetRevisionTypes();
}

TEST_F(WorkingWithRevisions, GetRevisionGroups)
{
    s_instance->GetRevisionGroups();
}

TEST_F(WorkingWithRevisions, RemoveCommentsInPdf)
{
    s_instance->RemoveCommentsInPdf();
}

TEST_F(WorkingWithRevisions, ShowRevisionsInBalloons)
{
    s_instance->ShowRevisionsInBalloons();
}

TEST_F(WorkingWithRevisions, GetRevisionGroupDetails)
{
    s_instance->GetRevisionGroupDetails();
}

TEST_F(WorkingWithRevisions, AccessRevisedVersion)
{
    s_instance->AccessRevisedVersion();
}

TEST_F(WorkingWithRevisions, MoveNodeInTrackedDocument)
{
    s_instance->MoveNodeInTrackedDocument();
}

TEST_F(WorkingWithRevisions, ShapeRevision)
{
    s_instance->ShapeRevision();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
