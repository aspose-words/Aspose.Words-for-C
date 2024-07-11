#include "Apply License.h"

using namespace Aspose::Words;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class ApplyLicense : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::ApplyLicense> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::ApplyLicense>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::ApplyLicense> ApplyLicense::s_instance;

TEST_F(ApplyLicense, ApplyLicenseFromFile)
{
    s_instance->ApplyLicenseFromFile();
}

TEST_F(ApplyLicense, ApplyLicenseFromStream)
{
    s_instance->ApplyLicenseFromStream();
}
#ifdef _WIN32
TEST_F(ApplyLicense, ApplyLicenseFromEmbeddedResourceWindows)
{
    s_instance->ApplyLicenseFromEmbeddedResourceWindows();
}
#elif __linux__
TEST_F(ApplyLicense, ApplyLicenseFromEmbeddedResourceLinux)
{
    s_instance->ApplyLicenseFromEmbeddedResourceLinux();
}
#endif

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
