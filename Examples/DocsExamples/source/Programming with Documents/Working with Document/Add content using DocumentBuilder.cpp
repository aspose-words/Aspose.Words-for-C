#include "Add content using DocumentBuilder.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Tables;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document { namespace gtest_test {

class AddContentUsingDocumentBuilder : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::AddContentUsingDocumentBuilder> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Document::AddContentUsingDocumentBuilder>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::AddContentUsingDocumentBuilder> AddContentUsingDocumentBuilder::s_instance;

TEST_F(AddContentUsingDocumentBuilder, CreateNewDocument)
{
    s_instance->CreateNewDocument();
}

TEST_F(AddContentUsingDocumentBuilder, DocumentBuilderInsertBookmark)
{
    s_instance->DocumentBuilderInsertBookmark();
}

TEST_F(AddContentUsingDocumentBuilder, BuildTable)
{
    s_instance->BuildTable();
}

TEST_F(AddContentUsingDocumentBuilder, InsertHorizontalRule)
{
    s_instance->InsertHorizontalRule();
}

TEST_F(AddContentUsingDocumentBuilder, HorizontalRuleFormat)
{
    s_instance->HorizontalRuleFormat_();
}

TEST_F(AddContentUsingDocumentBuilder, InsertBreak)
{
    s_instance->InsertBreak();
}

TEST_F(AddContentUsingDocumentBuilder, InsertTextInputFormField)
{
    s_instance->InsertTextInputFormField();
}

TEST_F(AddContentUsingDocumentBuilder, InsertCheckBoxFormField)
{
    s_instance->InsertCheckBoxFormField();
}

TEST_F(AddContentUsingDocumentBuilder, InsertComboBoxFormField)
{
    s_instance->InsertComboBoxFormField();
}

TEST_F(AddContentUsingDocumentBuilder, InsertHtml)
{
    s_instance->InsertHtml();
}

TEST_F(AddContentUsingDocumentBuilder, InsertHyperlink)
{
    s_instance->InsertHyperlink();
}

TEST_F(AddContentUsingDocumentBuilder, InsertTableOfContents)
{
    s_instance->InsertTableOfContents();
}

TEST_F(AddContentUsingDocumentBuilder, InsertInlineImage)
{
    s_instance->InsertInlineImage();
}

TEST_F(AddContentUsingDocumentBuilder, InsertFloatingImage)
{
    s_instance->InsertFloatingImage();
}

TEST_F(AddContentUsingDocumentBuilder, InsertParagraph)
{
    s_instance->InsertParagraph();
}

TEST_F(AddContentUsingDocumentBuilder, InsertTCField)
{
    s_instance->InsertTCField();
}

TEST_F(AddContentUsingDocumentBuilder, InsertTCFieldsAtText)
{
    s_instance->InsertTCFieldsAtText();
}

TEST_F(AddContentUsingDocumentBuilder, CursorPosition)
{
    s_instance->CursorPosition();
}

TEST_F(AddContentUsingDocumentBuilder, MoveToNode)
{
    s_instance->MoveToNode();
}

TEST_F(AddContentUsingDocumentBuilder, MoveToDocumentStartEnd)
{
    s_instance->MoveToDocumentStartEnd();
}

TEST_F(AddContentUsingDocumentBuilder, MoveToSection)
{
    s_instance->MoveToSection();
}

TEST_F(AddContentUsingDocumentBuilder, MoveToHeadersFooters)
{
    s_instance->MoveToHeadersFooters();
}

TEST_F(AddContentUsingDocumentBuilder, MoveToParagraph)
{
    s_instance->MoveToParagraph();
}

TEST_F(AddContentUsingDocumentBuilder, MoveToTableCell)
{
    s_instance->MoveToTableCell();
}

TEST_F(AddContentUsingDocumentBuilder, MoveToBookmarkEnd)
{
    s_instance->MoveToBookmarkEnd();
}

TEST_F(AddContentUsingDocumentBuilder, MoveToMergeField)
{
    s_instance->MoveToMergeField();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document::gtest_test
