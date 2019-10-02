#include "stdafx.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

// ExStart:FieldsHelper
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
// ExEnd:FieldsHelper