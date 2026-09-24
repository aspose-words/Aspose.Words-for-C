#include "ExRendering.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3011900444u, ::Aspose::Words::ApiExamples::ExRendering, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExRendering : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExRendering> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExRendering>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExRendering> ExRendering::s_instance;

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
