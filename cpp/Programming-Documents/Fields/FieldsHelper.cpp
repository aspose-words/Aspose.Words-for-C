#include "stdafx.h"

#include "FieldsHelper.h"

#include <system/enumerator_adapter.h>
#include <Model/Fields/Field.h>
#include <Model/Fields/FieldCollection.h>
#include <Model/Fields/FieldType.h>
#include <Model/Nodes/CompositeNode.h>
#include <Model/Text/Range.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void ConvertFieldsToStaticText(System::SharedPtr<CompositeNode> compositeNode, FieldType targetFieldType)
{
    for (System::SharedPtr<Field> field : System::IterateOver(compositeNode->get_Range()->get_Fields()))
    {
        if (field->get_Type() == targetFieldType)
        {
            field->Unlink();
        }
    }
}