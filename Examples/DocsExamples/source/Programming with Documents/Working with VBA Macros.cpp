#include "Working with VBA Macros.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Vba;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithVba : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithVba> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithVba>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithVba> WorkingWithVba::s_instance;

TEST_F(WorkingWithVba, CreateVbaProject)
{
    s_instance->CreateVbaProject();
}

TEST_F(WorkingWithVba, ReadVbaMacros)
{
    s_instance->ReadVbaMacros();
}

TEST_F(WorkingWithVba, ModifyVbaMacros)
{
    s_instance->ModifyVbaMacros();
}

TEST_F(WorkingWithVba, CloneVbaProject)
{
    s_instance->CloneVbaProject();
}

TEST_F(WorkingWithVba, CloneVbaModule)
{
    s_instance->CloneVbaModule();
}

TEST_F(WorkingWithVba, RemoveBrokenRef)
{
    s_instance->RemoveBrokenRef();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
