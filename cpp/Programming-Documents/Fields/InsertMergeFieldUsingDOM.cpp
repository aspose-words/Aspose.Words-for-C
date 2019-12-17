#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Fields/Fields/MailMerge/FieldMergeField.h>
#include <Model/Fields/FieldType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Text/Paragraph.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void InsertMergeFieldUsingDOM()
{
    std::cout << "InsertMergeFieldUsingDOM example started." << std::endl;
    // ExStart:InsertMergeFieldUsingDOM
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"in.doc");
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Get paragraph you want to append this merge field to
    System::SharedPtr<Paragraph> para = System::DynamicCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(1));

    // Move cursor to this paragraph
    builder->MoveTo(para);

    // We want to insert a merge field like this:
    // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }

    // Create instance of FieldMergeField class and lets build the above field code
    System::SharedPtr<FieldMergeField> field = System::DynamicCast<FieldMergeField>(builder->InsertField(FieldType::FieldMergeField, false));

    // { " MERGEFIELD Test1" }
    field->set_FieldName(u"Test1");

    // { " MERGEFIELD Test1 \\b Test2" }
    field->set_TextBefore(u"Test2");

    // { " MERGEFIELD Test1 \\b Test2 \\f Test3 }
    field->set_TextAfter(u"Test3");

    // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m" }
    field->set_IsMapped(true);

    // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
    field->set_IsVerticalFormatting(true);

    // Finally update this merge field
    field->Update();

    System::String outputPath = outputDataDir + u"InsertMergeFieldUsingDOM.doc";
    doc->Save(outputPath);

    // ExEnd:InsertMergeFieldUsingDOM
    std::cout << "Merge field using DOM inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertMergeFieldUsingDOM example finished." << std::endl << std::endl;
}