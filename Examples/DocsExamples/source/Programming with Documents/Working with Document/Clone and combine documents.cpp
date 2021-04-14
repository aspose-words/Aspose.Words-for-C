#include "Clone and combine documents.h"

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
using namespace Aspose::Words::Replacing;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document { namespace gtest_test {

class CloneAndCombineDocuments : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::CloneAndCombineDocuments> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Document::CloneAndCombineDocuments>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::CloneAndCombineDocuments> CloneAndCombineDocuments::s_instance;

TEST_F(CloneAndCombineDocuments, CloningDocument)
{
    s_instance->CloningDocument();
}

TEST_F(CloneAndCombineDocuments, InsertDocumentAtReplace)
{
    s_instance->InsertDocumentAtReplace();
}

TEST_F(CloneAndCombineDocuments, InsertDocumentAtBookmark)
{
    s_instance->InsertDocumentAtBookmark();
}

TEST_F(CloneAndCombineDocuments, InsertDocumentAtMailMerge)
{
    s_instance->InsertDocumentAtMailMerge();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document::gtest_test
