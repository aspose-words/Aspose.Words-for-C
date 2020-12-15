#include "ExSmartTag.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Markup/SmartTag.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;
namespace ApiExamples { namespace gtest_test {

class ExSmartTag : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExSmartTag> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExSmartTag>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExSmartTag> ExSmartTag::s_instance;

TEST_F(ExSmartTag, Create)
{
    s_instance->Create();
}

TEST_F(ExSmartTag, Properties)
{
    s_instance->Properties();
}

}} // namespace ApiExamples::gtest_test
