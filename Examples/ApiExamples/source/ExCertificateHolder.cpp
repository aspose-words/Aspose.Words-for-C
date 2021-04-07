#include "ExCertificateHolder.h"

namespace ApiExamples { namespace gtest_test {

class ExCertificateHolder : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExCertificateHolder> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExCertificateHolder>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExCertificateHolder> ExCertificateHolder::s_instance;

}} // namespace ApiExamples::gtest_test
