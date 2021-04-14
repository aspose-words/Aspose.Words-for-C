#include "Join and append documents.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document { namespace gtest_test {

class JoinAndAppendDocuments : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::JoinAndAppendDocuments> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Document::JoinAndAppendDocuments>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::JoinAndAppendDocuments> JoinAndAppendDocuments::s_instance;

TEST_F(JoinAndAppendDocuments, SimpleAppendDocument)
{
    s_instance->SimpleAppendDocument();
}

TEST_F(JoinAndAppendDocuments, AppendDocument)
{
    s_instance->AppendDocument();
}

TEST_F(JoinAndAppendDocuments, AppendDocumentToBlank)
{
    s_instance->AppendDocumentToBlank();
}

TEST_F(JoinAndAppendDocuments, AppendWithImportFormatOptions)
{
    s_instance->AppendWithImportFormatOptions();
}

TEST_F(JoinAndAppendDocuments, ConvertNumPageFields)
{
    s_instance->ConvertNumPageFields();
}

TEST_F(JoinAndAppendDocuments, DifferentPageSetup)
{
    s_instance->DifferentPageSetup();
}

TEST_F(JoinAndAppendDocuments, JoinContinuous)
{
    s_instance->JoinContinuous();
}

TEST_F(JoinAndAppendDocuments, JoinNewPage)
{
    s_instance->JoinNewPage();
}

TEST_F(JoinAndAppendDocuments, KeepSourceFormatting)
{
    s_instance->KeepSourceFormatting();
}

TEST_F(JoinAndAppendDocuments, KeepSourceTogether)
{
    s_instance->KeepSourceTogether();
}

TEST_F(JoinAndAppendDocuments, ListKeepSourceFormatting)
{
    s_instance->ListKeepSourceFormatting();
}

TEST_F(JoinAndAppendDocuments, ListUseDestinationStyles)
{
    s_instance->ListUseDestinationStyles();
}

TEST_F(JoinAndAppendDocuments, RestartPageNumbering)
{
    s_instance->RestartPageNumbering();
}

TEST_F(JoinAndAppendDocuments, UpdatePageLayout)
{
    s_instance->UpdatePageLayout();
}

TEST_F(JoinAndAppendDocuments, UseDestinationStyles)
{
    s_instance->UseDestinationStyles();
}

TEST_F(JoinAndAppendDocuments, SmartStyleBehavior)
{
    s_instance->SmartStyleBehavior();
}

TEST_F(JoinAndAppendDocuments, InsertDocumentWithBuilder)
{
    s_instance->InsertDocumentWithBuilder();
}

TEST_F(JoinAndAppendDocuments, KeepSourceNumbering)
{
    s_instance->KeepSourceNumbering();
}

TEST_F(JoinAndAppendDocuments, IgnoreTextBoxes)
{
    s_instance->IgnoreTextBoxes();
}

TEST_F(JoinAndAppendDocuments, IgnoreHeaderFooter)
{
    s_instance->IgnoreHeaderFooter();
}

TEST_F(JoinAndAppendDocuments, LinkHeadersFooters)
{
    s_instance->LinkHeadersFooters();
}

TEST_F(JoinAndAppendDocuments, RemoveSourceHeadersFooters)
{
    s_instance->RemoveSourceHeadersFooters();
}

TEST_F(JoinAndAppendDocuments, UnlinkHeadersFooters)
{
    s_instance->UnlinkHeadersFooters();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document::gtest_test
