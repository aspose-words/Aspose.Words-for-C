#include "stdafx.h"
#include "examples.h"

#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Fields/Field.h>
#include <Model/Fields/FieldType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void SpecifylocaleAtFieldlevel()
{
    std::cout << "SpecifylocaleAtFieldlevel example started." << std::endl;
    // ExStart:SpecifylocaleAtFieldlevel
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>();
    System::SharedPtr<Field> field = builder->InsertField(FieldType::FieldDate, true);
    field->set_LocaleId(1049);
    builder->get_Document()->Save(outputDataDir + u"SpecifylocaleAtFieldlevel.docx");
    // ExEnd:SpecifylocaleAtFieldlevel
    std::cout << "SpecifylocaleAtFieldlevel example finished." << std::endl << std::endl;
}
