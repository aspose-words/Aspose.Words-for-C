#include "Working with Footnote and Endnote.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Notes;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithFootnotes : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithFootnotes> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithFootnotes>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithFootnotes> WorkingWithFootnotes::s_instance;

TEST_F(WorkingWithFootnotes, SetFootNoteColumns)
{
    s_instance->SetFootNoteColumns();
}

TEST_F(WorkingWithFootnotes, SetFootnoteAndEndNotePosition)
{
    s_instance->SetFootnoteAndEndNotePosition();
}

TEST_F(WorkingWithFootnotes, SetEndnoteOptions)
{
    s_instance->SetEndnoteOptions();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
