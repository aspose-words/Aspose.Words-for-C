#include "Working with FormFields.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithFormFields : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithFormFields> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithFormFields>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithFormFields> WorkingWithFormFields::s_instance;

TEST_F(WorkingWithFormFields, InsertFormFields)
{
    s_instance->InsertFormFields();
}

TEST_F(WorkingWithFormFields, FormFieldsWorkWithProperties)
{
    s_instance->FormFieldsWorkWithProperties();
}

TEST_F(WorkingWithFormFields, FormFieldsGetFormFieldsCollection)
{
    s_instance->FormFieldsGetFormFieldsCollection();
}

TEST_F(WorkingWithFormFields, FormFieldsGetByName)
{
    s_instance->FormFieldsGetByName();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
