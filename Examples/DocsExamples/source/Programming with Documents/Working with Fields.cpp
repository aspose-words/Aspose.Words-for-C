#include "Working with Fields.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithFields : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithFields> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithFields>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithFields> WorkingWithFields::s_instance;

TEST_F(WorkingWithFields, ChangeFieldUpdateCultureSource)
{
    s_instance->ChangeFieldUpdateCultureSource();
}

TEST_F(WorkingWithFields, SpecifyLocaleAtFieldLevel)
{
    s_instance->SpecifyLocaleAtFieldLevel();
}

TEST_F(WorkingWithFields, ReplaceHyperlinks)
{
    s_instance->ReplaceHyperlinks();
}

TEST_F(WorkingWithFields, RenameMergeFields)
{
    s_instance->RenameMergeFields();
}

TEST_F(WorkingWithFields, RemoveField)
{
    s_instance->RemoveField();
}

TEST_F(WorkingWithFields, InsertTOAFieldWithoutDocumentBuilder)
{
    s_instance->InsertTOAFieldWithoutDocumentBuilder();
}

TEST_F(WorkingWithFields, InsertNestedFields)
{
    s_instance->InsertNestedFields();
}

TEST_F(WorkingWithFields, InsertMergeFieldUsingDOM)
{
    s_instance->InsertMergeFieldUsingDOM();
}

TEST_F(WorkingWithFields, InsertMailMergeAddressBlockFieldUsingDOM)
{
    s_instance->InsertMailMergeAddressBlockFieldUsingDOM();
}

TEST_F(WorkingWithFields, InsertFieldIncludeTextWithoutDocumentBuilder)
{
    s_instance->InsertFieldIncludeTextWithoutDocumentBuilder();
}

TEST_F(WorkingWithFields, InsertFieldNone)
{
    s_instance->InsertFieldNone();
}

TEST_F(WorkingWithFields, InsertField)
{
    s_instance->InsertField();
}

TEST_F(WorkingWithFields, InsertAuthorField)
{
    s_instance->InsertAuthorField();
}

TEST_F(WorkingWithFields, InsertASKFieldWithOutDocumentBuilder)
{
    s_instance->InsertASKFieldWithOutDocumentBuilder();
}

TEST_F(WorkingWithFields, InsertAdvanceFieldWithOutDocumentBuilder)
{
    s_instance->InsertAdvanceFieldWithOutDocumentBuilder();
}

TEST_F(WorkingWithFields, GetMailMergeFieldNames)
{
    s_instance->GetMailMergeFieldNames();
}

TEST_F(WorkingWithFields, MappedDataFields)
{
    s_instance->MappedDataFields();
}

TEST_F(WorkingWithFields, DeleteFields)
{
    s_instance->DeleteFields();
}

TEST_F(WorkingWithFields, FieldUpdateCulture)
{
    s_instance->FieldUpdateCulture();
}

TEST_F(WorkingWithFields, FieldDisplayResults)
{
    s_instance->FieldDisplayResults();
}

TEST_F(WorkingWithFields, EvaluateIFCondition)
{
    s_instance->EvaluateIFCondition();
}

TEST_F(WorkingWithFields, ConvertFieldsInParagraph)
{
    s_instance->ConvertFieldsInParagraph();
}

TEST_F(WorkingWithFields, ConvertFieldsInDocument)
{
    s_instance->ConvertFieldsInDocument();
}

TEST_F(WorkingWithFields, ConvertFieldsInBody)
{
    s_instance->ConvertFieldsInBody();
}

TEST_F(WorkingWithFields, ChangeLocale)
{
    s_instance->ChangeLocale();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
