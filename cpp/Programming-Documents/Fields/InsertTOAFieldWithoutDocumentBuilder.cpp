#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Fields/Fields/IndexAndTables/FieldTA.h>
#include <Model/Fields/Fields/IndexAndTables/FieldToa.h>
#include <Model/Fields/FieldType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/Section.h>
#include <Model/Text/Paragraph.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void InsertTOAFieldWithoutDocumentBuilder()
{
    std::cout << "InsertTOAFieldWithoutDocumentBuilder example started." << std::endl;
    // ExStart:InsertTOAFieldWithoutDocumentBuilder
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"in.doc");
    // Get paragraph you want to append this TOA field to
    System::SharedPtr<Paragraph> para = System::DynamicCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(1));

    // We want to insert TA and TOA fields like this:
    // { TA  \c 1 \l "Value 0" }
    // { TOA  \c 1 }

    // Create instance of FieldAsk class and lets build the above field code
    System::SharedPtr<FieldTA> fieldTA = System::DynamicCast<FieldTA>(para->AppendField(FieldType::FieldTOAEntry, false));
    fieldTA->set_EntryCategory(u"1");
    fieldTA->set_LongCitation(u"Value 0");

    doc->get_FirstSection()->get_Body()->AppendChild(para);

    para = System::MakeObject<Paragraph>(doc);

    // Create instance of FieldToa class
    System::SharedPtr<FieldToa> fieldToa = System::DynamicCast<FieldToa>(para->AppendField(FieldType::FieldTOA, false));
    fieldToa->set_EntryCategory(u"1");
    doc->get_FirstSection()->get_Body()->AppendChild(para);

    // Finally update this TOA field
    fieldToa->Update();

    System::String outputPath = outputDataDir + u"InsertTOAFieldWithoutDocumentBuilder.doc";
    doc->Save(outputPath);

    // ExEnd:InsertTOAFieldWithoutDocumentBuilder
    std::cout << "TOA field without using document builder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertTOAFieldWithoutDocumentBuilder example finished." << std::endl << std::endl;
}