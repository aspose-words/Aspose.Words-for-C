#include "stdafx.h"
#include "examples.h"

#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Fields/Field.h>
#include <Model/Fields/FieldType.h>
#include <cstdint>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void SpecifylocaleAtFieldlevel()
{
    std::cout << "SpecifylocaleAtFieldlevel example started." << std::endl;
    // ExStart:SpecifylocaleAtFieldlevel
    System::String dataDir = GetDataDir_WorkingWithFields();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>();
    System::SharedPtr<Field> field = builder->InsertField(FieldType::FieldDate, true);
    field->set_LocaleId(1049);
    builder->get_Document()->Save(dataDir + GetOutputFilePath(u"SpecifylocaleAtFieldlevel.docx"));
    // ExEnd:SpecifylocaleAtFieldlevel
    std::cout << "SpecifylocaleAtFieldlevel example finished." << std::endl << std::endl;
}
