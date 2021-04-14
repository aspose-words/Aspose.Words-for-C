#include "Working with List.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Saving;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithList : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithList> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithList>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithList> WorkingWithList::s_instance;

TEST_F(WorkingWithList, RestartListAtEachSection)
{
    s_instance->RestartListAtEachSection();
}

TEST_F(WorkingWithList, SpecifyListLevel)
{
    s_instance->SpecifyListLevel();
}

TEST_F(WorkingWithList, RestartListNumber)
{
    s_instance->RestartListNumber();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
