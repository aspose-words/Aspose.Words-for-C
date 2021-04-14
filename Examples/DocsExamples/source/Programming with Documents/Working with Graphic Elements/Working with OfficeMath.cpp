#include "Working with OfficeMath.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Math;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Graphic_Elements { namespace gtest_test {

class WorkingWithOfficeMath : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithOfficeMath> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithOfficeMath>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithOfficeMath> WorkingWithOfficeMath::s_instance;

TEST_F(WorkingWithOfficeMath, MathEquations)
{
    s_instance->MathEquations();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::gtest_test
