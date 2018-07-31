#include "stdafx.h"
#include "examples.h"

#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Fields/Field.h>
#include <Model/Document/DocumentBuilder.h>
#include <cstdint>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void SpecifylocaleAtFieldlevel()
{
    // ExStart:SpecifylocaleAtFieldlevel
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>();
    System::SharedPtr<Field> field = builder->InsertField(u"=1", nullptr);
    field->set_LocaleId(1027);
    // ExEnd:SpecifylocaleAtFieldlevel
}
