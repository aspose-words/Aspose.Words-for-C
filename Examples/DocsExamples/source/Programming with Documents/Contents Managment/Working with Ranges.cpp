#include "Working with Ranges.h"

using namespace Aspose::Words;
namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Managment { namespace gtest_test {

class WorkingWithRanges : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Managment::WorkingWithRanges> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Contents_Managment::WorkingWithRanges>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Managment::WorkingWithRanges> WorkingWithRanges::s_instance;

TEST_F(WorkingWithRanges, RangesDeleteText)
{
    s_instance->RangesDeleteText();
}

TEST_F(WorkingWithRanges, RangesGetText)
{
    s_instance->RangesGetText();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Contents_Managment::gtest_test
