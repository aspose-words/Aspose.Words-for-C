#include "Working with AI.h"

using namespace Aspose::Words;
using namespace Aspose::Words::AI;
namespace DocsExamples { namespace AI_powered_Features { namespace gtest_test {

class WorkingWithAI : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::AI_powered_Features::WorkingWithAI> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::AI_powered_Features::WorkingWithAI>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::AI_powered_Features::WorkingWithAI> WorkingWithAI::s_instance;

// This test should be run manually to manage API requests amount.
TEST_F(WorkingWithAI, DISABLED_AiSummarize)
{
    s_instance->AiSummarize();
}

// This test should be run manually to manage API requests amount.
TEST_F(WorkingWithAI, DISABLED_AiTranslate)
{
    s_instance->AiTranslate();
}

// This test should be run manually to manage API requests amount.
TEST_F(WorkingWithAI, DISABLED_AiGrammar)
{
    s_instance->AiGrammar();
}

}}} // namespace DocsExamples::AI_powered_Features::gtest_test
