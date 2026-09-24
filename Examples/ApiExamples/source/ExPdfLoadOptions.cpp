#include "ExPdfLoadOptions.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(951908822u, ::Aspose::Words::ApiExamples::ExPdfLoadOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExPdfLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExPdfLoadOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExPdfLoadOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExPdfLoadOptions> ExPdfLoadOptions::s_instance;

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
