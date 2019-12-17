#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/EquationsAndFormulas/FieldAdvance.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void InsertAdvanceFieldWithoutDocumentBuilder()
{
    std::cout << "InsertAdvanceFieldWithoutDocumentBuilder example started." << std::endl;
    // ExStart:InsertAdvanceFieldWithoutDocumentBuilder
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"in.doc");
    // Get paragraph you want to append this Advance field to
    System::SharedPtr<Paragraph> para = System::DynamicCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(1));

    // We want to insert an Advance field like this:
    // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }

    // Create instance of FieldAdvance class and lets build the above field code
    System::SharedPtr<FieldAdvance> field = System::DynamicCast<FieldAdvance>(para->AppendField(FieldType::FieldAdvance, false));


    // { ADVANCE \\d 10 " }
    field->set_DownOffset(u"10");

    // { ADVANCE \\d 10 \\l 10 }
    field->set_LeftOffset(u"10");

    // { ADVANCE \\d 10 \\l 10 \\r -3.3 }
    field->set_RightOffset(u"-3.3");

    // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 }
    field->set_UpOffset(u"0");

    // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 }
    field->set_HorizontalPosition(u"100");

    // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }
    field->set_VerticalPosition(u"100");

    // Finally update this Advance field
    field->Update();

    System::String outputPath = outputDataDir + u"InsertAdvanceFieldWithoutDocumentBuilder.doc";
    doc->Save(outputPath);

    // ExEnd:InsertAdvanceFieldWithoutDocumentBuilder
    std::cout << "Advance field without using document builder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertAdvanceFieldWithoutDocumentBuilder example finished." << std::endl << std::endl;
}