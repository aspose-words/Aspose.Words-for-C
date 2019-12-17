#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Fields/Nodes/FieldStart.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Sections/Section.h>
#include <Model/Text/Paragraph.h>

#include "Common.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void ExtractContentUsingField()
{
    std::cout << "ExtractContentUsingField example started." << std::endl;
    // ExStart:ExtractContentUsingField
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");

    // Use a document builder to retrieve the field start of a merge field.
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
    // We could also get FieldStarts of a field using GetChildNode method as in the other examples.
    builder->MoveToMergeField(u"Fullname", false, false);

    // The builder cursor should be positioned at the start of the field.
    System::SharedPtr<FieldStart> startField = System::DynamicCast<FieldStart>(builder->get_CurrentNode());
    System::SharedPtr<Paragraph> endPara = System::DynamicCast<Paragraph>(doc->get_FirstSection()->GetChild(NodeType::Paragraph, 5, true));

    // Extract the content between these nodes in the document. Don't include these markers in the extraction.
    std::vector<System::SharedPtr<Node>> extractedNodes = ExtractContent(startField, endPara, false);

    // Insert the content into a new separate document and save it to disk.
    System::SharedPtr<Document> dstDoc = GenerateDocument(doc, extractedNodes);
    System::String outputPath = outputDataDir + u"ExtractContentUsingField.doc";
    dstDoc->Save(outputPath);
    // ExEnd:ExtractContentUsingField
    std::cout << "Extracted content using the Field successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ExtractContentUsingField example finished." << std::endl << std::endl;
}