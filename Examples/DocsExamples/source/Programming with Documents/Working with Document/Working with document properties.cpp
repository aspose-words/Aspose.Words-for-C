#include "Working with document properties.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Properties;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document { namespace gtest_test {

class DocumentPropertiesAndVariables : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::DocumentPropertiesAndVariables> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Document::DocumentPropertiesAndVariables>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Document::DocumentPropertiesAndVariables> DocumentPropertiesAndVariables::s_instance;

TEST_F(DocumentPropertiesAndVariables, GetVariables)
{
    s_instance->GetVariables();
}

TEST_F(DocumentPropertiesAndVariables, EnumerateProperties)
{
    s_instance->EnumerateProperties();
}

TEST_F(DocumentPropertiesAndVariables, AddCustomDocumentProperties)
{
    s_instance->AddCustomDocumentProperties();
}

TEST_F(DocumentPropertiesAndVariables, RemoveCustomDocumentProperties)
{
    s_instance->RemoveCustomDocumentProperties();
}

TEST_F(DocumentPropertiesAndVariables, RemovePersonalInformation)
{
    s_instance->RemovePersonalInformation();
}

TEST_F(DocumentPropertiesAndVariables, ConfiguringLinkToContent)
{
    s_instance->ConfiguringLinkToContent();
}

TEST_F(DocumentPropertiesAndVariables, ConvertBetweenMeasurementUnits)
{
    s_instance->ConvertBetweenMeasurementUnits();
}

TEST_F(DocumentPropertiesAndVariables, UseControlCharacters)
{
    s_instance->UseControlCharacters();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document::gtest_test
