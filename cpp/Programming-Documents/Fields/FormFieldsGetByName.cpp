#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void FormFieldsGetByName()
{
    std::cout << "FormFieldsGetByName example started." << std::endl;
    // ExStart:FormFieldsGetByName
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"FormFields.doc");
    System::SharedPtr<FormFieldCollection> documentFormFields = doc->get_Range()->get_FormFields();

    System::SharedPtr<FormField> formField1 = documentFormFields->idx_get(3);
    System::SharedPtr<FormField> formField2 = documentFormFields->idx_get(u"Text2");
    // ExEnd:FormFieldsGetByName
    std::cout << formField2->get_Name().ToUtf8String() << " field have following text " << formField2->GetText().ToUtf8String() << std::endl;
    std::cout << "FormFieldsGetByName example finished." << std::endl << std::endl;
}
