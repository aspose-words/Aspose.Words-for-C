#include "Extract content.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Tables;
namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Management { namespace gtest_test {

class ExtractContent : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Management::ExtractContent> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Contents_Management::ExtractContent>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Management::ExtractContent> ExtractContent::s_instance;

TEST_F(ExtractContent, ExtractContentBetweenBlockLevelNodes)
{
    s_instance->ExtractContentBetweenBlockLevelNodes();
}

TEST_F(ExtractContent, ExtractContentBetweenBookmark)
{
    s_instance->ExtractContentBetweenBookmark();
}

TEST_F(ExtractContent, ExtractContentBetweenCommentRange)
{
    s_instance->ExtractContentBetweenCommentRange();
}

TEST_F(ExtractContent, ExtractContentBetweenParagraphs)
{
    s_instance->ExtractContentBetweenParagraphs();
}

TEST_F(ExtractContent, ExtractContentBetweenParagraphStyles)
{
    s_instance->ExtractContentBetweenParagraphStyles();
}

TEST_F(ExtractContent, ExtractContentBetweenRuns)
{
    s_instance->ExtractContentBetweenRuns();
}

TEST_F(ExtractContent, ExtractContentUsingDocumentVisitor)
{
    s_instance->ExtractContentUsingDocumentVisitor();
}

TEST_F(ExtractContent, ExtractContentUsingField)
{
    s_instance->ExtractContentUsingField();
}

TEST_F(ExtractContent, ExtractTableOfContents)
{
    s_instance->ExtractTableOfContents();
}

TEST_F(ExtractContent, ExtractTextOnly)
{
    s_instance->ExtractTextOnly();
}

TEST_F(ExtractContent, ExtractContentBasedOnStyles)
{
    s_instance->ExtractContentBasedOnStyles();
}

TEST_F(ExtractContent, ExtractPrintText)
{
    s_instance->ExtractPrintText();
}

TEST_F(ExtractContent, ExtractImagesToFiles)
{
    s_instance->ExtractImagesToFiles();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Contents_Management::gtest_test
