#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Fields/Fields/EquationsAndFormulas/FieldAdvance.h>
#include <Model/Fields/FieldType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Text/Paragraph.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void InsertAdvanceFieldWithoutDocumentBuilder()
{
    std::cout << "InsertAdvanceFieldWithoutDocumentBuilder example started." << std::endl;
    // ExStart:InsertAdvanceFieldWithoutDocumentBuilder
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithFields();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"in.doc");
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

    System::String outputPath = dataDir + GetOutputFilePath(u"InsertAdvanceFieldWithoutDocumentBuilder.doc");
    doc->Save(outputPath);

    // ExEnd:InsertAdvanceFieldWithoutDocumentBuilder
    std::cout << "Advance field without using document builder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertAdvanceFieldWithoutDocumentBuilder example finished." << std::endl << std::endl;
}