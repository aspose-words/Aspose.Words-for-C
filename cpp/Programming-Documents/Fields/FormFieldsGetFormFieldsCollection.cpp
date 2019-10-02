#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void FormFieldsGetFormFieldsCollection()
{
    std::cout << "FormFieldsGetFormFieldsCollection example started." << std::endl;
    // ExStart:FormFieldsGetFormFieldsCollection
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"FormFields.doc");
    System::SharedPtr<FormFieldCollection> formFields = doc->get_Range()->get_FormFields();

    // ExEnd:FormFieldsGetFormFieldsCollection
    std::cout << "Document have " << formFields->get_Count() << " form fields." << std::endl;
    std::cout << "FormFieldsGetFormFieldsCollection example finished." << std::endl << std::endl;
}
