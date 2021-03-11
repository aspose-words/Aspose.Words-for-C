#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>

using namespace Aspose::Words;

void FieldDisplayResults()
{
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();

    //ExStart:FieldDisplayResults
    //ExStart:UpdateDocFields
    auto document = System::MakeObject<Document>(inputDataDir + u"Various fields.docx");

    document->UpdateFields();
    //ExEnd:UpdateDocFields

    for(auto field : System::IterateOver(document->get_Range()->get_Fields()))
    {
        std::cout << field->get_DisplayResult().ToUtf8String();
    }
    //ExEnd:FieldDisplayResults
}
