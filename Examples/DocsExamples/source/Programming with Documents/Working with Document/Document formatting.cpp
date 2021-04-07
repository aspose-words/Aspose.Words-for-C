#include "Document formatting.h"

using namespace Aspose::Words;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document { namespace gtest_test {

class DocumentFormatting : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::DocumentFormatting> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Document::DocumentFormatting>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::DocumentFormatting> DocumentFormatting::s_instance;

TEST_F(DocumentFormatting, SpaceBetweenAsianAndLatinText)
{
    s_instance->SpaceBetweenAsianAndLatinText();
}

TEST_F(DocumentFormatting, AsianTypographyLineBreakGroup)
{
    s_instance->AsianTypographyLineBreakGroup();
}

TEST_F(DocumentFormatting, ParagraphFormatting)
{
    s_instance->ParagraphFormatting();
}

TEST_F(DocumentFormatting, MultilevelListFormatting)
{
    s_instance->MultilevelListFormatting();
}

TEST_F(DocumentFormatting, ApplyParagraphStyle)
{
    s_instance->ApplyParagraphStyle();
}

TEST_F(DocumentFormatting, ApplyBordersAndShadingToParagraph)
{
    s_instance->ApplyBordersAndShadingToParagraph();
}

TEST_F(DocumentFormatting, ChangeAsianParagraphSpacingAndIndents)
{
    s_instance->ChangeAsianParagraphSpacingAndIndents();
}

TEST_F(DocumentFormatting, SnapToGrid)
{
    s_instance->SnapToGrid();
}

TEST_F(DocumentFormatting, GetParagraphStyleSeparator)
{
    s_instance->GetParagraphStyleSeparator();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document::gtest_test
