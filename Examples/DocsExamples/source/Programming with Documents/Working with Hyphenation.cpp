#include "Working with Hyphenation.h"

using namespace Aspose::Words;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithHyphenation : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithHyphenation> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithHyphenation>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithHyphenation> WorkingWithHyphenation::s_instance;

TEST_F(WorkingWithHyphenation, HyphenateWordsOfLanguages)
{
    s_instance->HyphenateWordsOfLanguages();
}

TEST_F(WorkingWithHyphenation, LoadHyphenationDictionaryForLanguage)
{
    s_instance->LoadHyphenationDictionaryForLanguage();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
