#include "Compare documents.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Comparing;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document { namespace gtest_test {

class CompareDocument : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::CompareDocument> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Document::CompareDocument>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::CompareDocument> CompareDocument::s_instance;

TEST_F(CompareDocument, CompareForEqual)
{
    s_instance->CompareForEqual();
}

TEST_F(CompareDocument, CompareOptions)
{
    s_instance->CompareOptions_();
}

TEST_F(CompareDocument, ComparisonTarget)
{
    s_instance->ComparisonTarget();
}

TEST_F(CompareDocument, ComparisonGranularity)
{
    s_instance->ComparisonGranularity();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document::gtest_test
