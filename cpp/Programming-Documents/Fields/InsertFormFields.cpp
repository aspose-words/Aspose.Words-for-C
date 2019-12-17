#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void InsertFormFields()
{
    std::cout << "InsertFormFields example started." << std::endl;
    // ExStart:InsertFormFields
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    System::ArrayPtr<System::String> items = System::MakeArray<System::String>({u"One", u"Two", u"Three"});
    builder->InsertComboBox(u"DropDown", items, 0);
    // ExEnd:InsertFormFields
    std::cout << "Form fields inserted successfully." << std::endl;
    std::cout << "InsertFormFields example finished." << std::endl << std::endl;
}
