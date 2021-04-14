#include "ExMisc.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
namespace ApiExamples { namespace gtest_test {

class ExMisc : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExMisc> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExMisc>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExMisc> ExMisc::s_instance;

TEST_F(ExMisc, TrackingChanges)
{
    s_instance->TrackingChanges();
}

TEST_F(ExMisc, MoveNodeWithinATrackedDocument)
{
    s_instance->MoveNodeWithinATrackedDocument();
}

TEST_F(ExMisc, ApplyDifferentPropertiesWithRevisions)
{
    s_instance->ApplyDifferentPropertiesWithRevisions();
}

}} // namespace ApiExamples::gtest_test
