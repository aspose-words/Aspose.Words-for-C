#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Text/Range.h>
#include <Model/Fields/FormFields/FormFieldCollection.h>
#include <Model/Document/Document.h>
#include <cstdint>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void FormFieldsGetFormFieldsCollection()
{
    // ExStart:FormFieldsGetFormFieldsCollection
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithFields();
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"FormFields.doc");
    System::SharedPtr<FormFieldCollection> formFields = doc->get_Range()->get_FormFields();
    
    // ExEnd:FormFieldsGetFormFieldsCollection
    std::cout << "\nDocument have " << formFields->get_Count() << " form fields.\n";
}
