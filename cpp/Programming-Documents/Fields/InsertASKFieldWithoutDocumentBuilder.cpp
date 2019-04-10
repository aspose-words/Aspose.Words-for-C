#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Fields/Fields/LinksAndReferences/FieldAsk.h>
#include <Model/Fields/FieldType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Text/Paragraph.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void InsertASKFieldWithoutDocumentBuilder()
{
    std::cout << "InsertASKFieldWithoutDocumentBuilder example started." << std::endl;
    // ExStart:InsertASKFieldWithoutDocumentBuilder
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithFields();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"in.doc");
    // Get paragraph you want to append this Ask field to
    System::SharedPtr<Paragraph> para = System::DynamicCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(1));

    // We want to insert an Ask field like this:
    // { ASK \"Test 1\" Test2 \\d Test3 \\o }

    // Create instance of FieldAsk class and lets build the above field code
    System::SharedPtr<FieldAsk> field = System::DynamicCast<FieldAsk>(para->AppendField(FieldType::FieldAsk, false));

    // { ASK \"Test 1\" " }
    field->set_BookmarkName(u"Test 1");

    // { ASK \"Test 1\" Test2 }
    field->set_PromptText(u"Test2");

    // { ASK \"Test 1\" Test2 \\d Test3 }
    field->set_DefaultResponse(u"Test3");

    // { ASK \"Test 1\" Test2 \\d Test3 \\o }
    field->set_PromptOnceOnMailMerge(true);

    // Finally update this Ask field
    field->Update();

    System::String outputPath = dataDir + GetOutputFilePath(u"InsertASKFieldWithoutDocumentBuilder.doc");
    doc->Save(outputPath);

    // ExEnd:InsertASKFieldWithoutDocumentBuilder
    std::cout << "ASK field without using document builder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertASKFieldWithoutDocumentBuilder example finished." << std::endl << std::endl;
}