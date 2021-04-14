#include "Working with document options and settings.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Settings;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document { namespace gtest_test {

class WorkingWithDocumentOptionsAndSettings : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::WorkingWithDocumentOptionsAndSettings> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Document::WorkingWithDocumentOptionsAndSettings>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::WorkingWithDocumentOptionsAndSettings>
    WorkingWithDocumentOptionsAndSettings::s_instance;

TEST_F(WorkingWithDocumentOptionsAndSettings, OptimizeForMsWord)
{
    s_instance->OptimizeForMsWord();
}

TEST_F(WorkingWithDocumentOptionsAndSettings, ShowGrammaticalAndSpellingErrors)
{
    s_instance->ShowGrammaticalAndSpellingErrors();
}

TEST_F(WorkingWithDocumentOptionsAndSettings, CleanupUnusedStylesAndLists)
{
    s_instance->CleanupUnusedStylesAndLists();
}

TEST_F(WorkingWithDocumentOptionsAndSettings, CleanupDuplicateStyle)
{
    s_instance->CleanupDuplicateStyle();
}

TEST_F(WorkingWithDocumentOptionsAndSettings, ViewOptions)
{
    s_instance->ViewOptions();
}

TEST_F(WorkingWithDocumentOptionsAndSettings, DocumentPageSetup)
{
    s_instance->DocumentPageSetup();
}

TEST_F(WorkingWithDocumentOptionsAndSettings, AddJapaneseAsEditingLanguages)
{
    s_instance->AddJapaneseAsEditingLanguages();
}

TEST_F(WorkingWithDocumentOptionsAndSettings, SetRussianAsDefaultEditingLanguage)
{
    s_instance->SetRussianAsDefaultEditingLanguage();
}

TEST_F(WorkingWithDocumentOptionsAndSettings, SetPageSetupAndSectionFormatting)
{
    s_instance->SetPageSetupAndSectionFormatting();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document::gtest_test
