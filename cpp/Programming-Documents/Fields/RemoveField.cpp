#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Text/Range.h>
#include <Model/Nodes/Node.h>
#include <Model/Fields/FieldCollection.h>
#include <Model/Fields/Field.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void RemoveField()
{
    std::cout << "RemoveField example started." << std::endl;
    // ExStart:RemoveField
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Field.RemoveField.doc");

    System::SharedPtr<Field> field = doc->get_Range()->get_Fields()->idx_get(0);
    // Calling this method completely removes the field from the document.
    field->Remove();
    // ExEnd:RemoveField
    std::cout << "Removed field from the document successfully." << std::endl;
    std::cout << "RemoveField example finished." << std::endl << std::endl;
}
