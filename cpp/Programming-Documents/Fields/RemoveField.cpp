#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void RemoveField()
{
    std::cout << "RemoveField example started." << std::endl;
    // ExStart:RemoveField
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Field.RemoveField.doc");

    System::SharedPtr<Field> field = doc->get_Range()->get_Fields()->idx_get(0);
    // Calling this method completely removes the field from the document.
    field->Remove();
    // ExEnd:RemoveField
    std::cout << "Removed field from the document successfully." << std::endl;
    std::cout << "RemoveField example finished." << std::endl << std::endl;
}
