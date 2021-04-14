#include "Working with Textboxes.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithTextboxes : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithTextboxes> s_instance;

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithTextboxes>();
    };

    static void TearDownTestCase()
    {
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithTextboxes> WorkingWithTextboxes::s_instance;

TEST_F(WorkingWithTextboxes, CreateALink)
{
    s_instance->CreateALink();
}

TEST_F(WorkingWithTextboxes, CheckSequence)
{
    s_instance->CheckSequence();
}

TEST_F(WorkingWithTextboxes, BreakALink)
{
    s_instance->BreakALink();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
