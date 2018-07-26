#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/array.h>
#include <Model/Fields/FormFields/FormField.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void InsertFormFields()
{
    // ExStart:InsertFormFields
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    System::ArrayPtr<System::String> items = System::MakeArray<System::String>({u"One", u"Two", u"Three"});
    builder->InsertComboBox(u"DropDown", items, 0);
    // ExEnd:InsertFormFields
    std::cout << "\nForm fields inserted successfully.\n";
}
