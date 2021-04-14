#include "Working with WebExtension.h"

using namespace Aspose::Words;
using namespace Aspose::Words::WebExtensions;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document { namespace gtest_test {

class WorkingWithWebExtension : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::WorkingWithWebExtension> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Document::WorkingWithWebExtension>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::WorkingWithWebExtension> WorkingWithWebExtension::s_instance;

TEST_F(WorkingWithWebExtension, UsingWebExtensionTaskPanes)
{
    s_instance->UsingWebExtensionTaskPanes();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document::gtest_test
