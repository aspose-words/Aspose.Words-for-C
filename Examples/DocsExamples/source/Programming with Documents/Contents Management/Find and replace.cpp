#include "Find and replace.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Tables;
namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Management { namespace gtest_test {

class FindAndReplace : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Management::FindAndReplace> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Contents_Management::FindAndReplace>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Management::FindAndReplace> FindAndReplace::s_instance;

TEST_F(FindAndReplace, SimpleFindReplace)
{
    s_instance->SimpleFindReplace();
}

TEST_F(FindAndReplace, FindAndHighlight)
{
    s_instance->FindAndHighlight();
}

TEST_F(FindAndReplace, MetaCharactersInSearchPattern)
{
    s_instance->MetaCharactersInSearchPattern();
}

TEST_F(FindAndReplace, ReplaceTextContainingMetaCharacters)
{
    s_instance->ReplaceTextContainingMetaCharacters();
}

TEST_F(FindAndReplace, IgnoreTextInsideFields)
{
    s_instance->IgnoreTextInsideFields();
}

TEST_F(FindAndReplace, IgnoreTextInsideDeleteRevisions)
{
    s_instance->IgnoreTextInsideDeleteRevisions();
}

TEST_F(FindAndReplace, IgnoreTextInsideInsertRevisions)
{
    s_instance->IgnoreTextInsideInsertRevisions();
}

TEST_F(FindAndReplace, ReplaceHtmlTextWithMetaCharacters)
{
    s_instance->ReplaceHtmlTextWithMetaCharacters();
}

TEST_F(FindAndReplace, ReplaceTextInFooter)
{
    s_instance->ReplaceTextInFooter();
}

TEST_F(FindAndReplace, ShowChangesForHeaderAndFooterOrders)
{
    s_instance->ShowChangesForHeaderAndFooterOrders();
}

TEST_F(FindAndReplace, ReplaceTextWithField)
{
    s_instance->ReplaceTextWithField();
}

TEST_F(FindAndReplace, ReplaceWithEvaluator)
{
    s_instance->ReplaceWithEvaluator();
}

TEST_F(FindAndReplace, ReplaceWithHtml)
{
    s_instance->ReplaceWithHtml();
}

TEST_F(FindAndReplace, ReplaceWithRegex)
{
    s_instance->ReplaceWithRegex();
}

TEST_F(FindAndReplace, RecognizeAndSubstitutionsWithinReplacementPatterns)
{
    s_instance->RecognizeAndSubstitutionsWithinReplacementPatterns();
}

TEST_F(FindAndReplace, ReplaceWithString)
{
    s_instance->ReplaceWithString();
}

TEST_F(FindAndReplace, UsingLegacyOrder)
{
    s_instance->UsingLegacyOrder();
}

TEST_F(FindAndReplace, ReplaceTextInTable)
{
    s_instance->ReplaceTextInTable();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Contents_Management::gtest_test
