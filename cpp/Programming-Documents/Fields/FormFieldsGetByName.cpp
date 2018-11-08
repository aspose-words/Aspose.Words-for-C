#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Text/Range.h>
#include <Model/Fields/FormFields/FormFieldCollection.h>
#include <Model/Fields/FormFields/FormField.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void FormFieldsGetByName()
{
    std::cout << "FormFieldsGetByName example started." << std::endl;
    // ExStart:FormFieldsGetByName
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"FormFields.doc");
    System::SharedPtr<FormFieldCollection> documentFormFields = doc->get_Range()->get_FormFields();

    System::SharedPtr<FormField> formField1 = documentFormFields->idx_get(3);
    System::SharedPtr<FormField> formField2 = documentFormFields->idx_get(u"Text2");
    // ExEnd:FormFieldsGetByName
    std::cout << formField2->get_Name().ToUtf8String() << " field have following text " << formField2->GetText().ToUtf8String() << std::endl;
    std::cout << "FormFieldsGetByName example finished." << std::endl << std::endl;
}
