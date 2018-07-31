#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Text/Range.h>
#include <Model/Fields/FormFields/FormFieldCollection.h>
#include <Model/Fields/FormFields/FormField.h>
#include <Model/Fields/FieldType.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void FormFieldsWorkWithProperties()
{
    // ExStart:FormFieldsWorkWithProperties
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithFields();
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"FormFields.doc");
    System::SharedPtr<FormField> formField = doc->get_Range()->get_FormFields()->idx_get(3);
    
    if (formField->get_Type() == FieldType::FieldFormTextInput)
    {
        formField->set_Result(System::String(u"My name is ") + formField->get_Name());
    }
    // ExEnd:FormFieldsWorkWithProperties
}
