#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void FormFieldsWorkWithProperties()
{
    std::cout << "FormFieldsWorkWithProperties example started." << std::endl;
    // ExStart:FormFieldsWorkWithProperties
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"FormFields.doc");
    System::SharedPtr<FormField> formField = doc->get_Range()->get_FormFields()->idx_get(3);

    if (formField->get_Type() == FieldType::FieldFormTextInput)
    {
        formField->set_Result(System::String(u"My name is ") + formField->get_Name());
    }
    // ExEnd:FormFieldsWorkWithProperties
    std::cout << "FormFieldsWorkWithProperties example finished." << std::endl << std::endl;
}
