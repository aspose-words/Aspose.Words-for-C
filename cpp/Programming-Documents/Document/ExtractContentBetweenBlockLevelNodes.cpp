#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "Common.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

void ExtractContentBetweenBlockLevelNodes()
{
    std::cout << "ExtractContentBetweenBlockLevelNodes example started." << std::endl;
    // ExStart:ExtractContentBetweenBlockLevelNodes
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");

    System::SharedPtr<Paragraph> startPara = System::DynamicCast<Paragraph>(doc->get_LastSection()->GetChild(NodeType::Paragraph, 2, true));
    System::SharedPtr<Table> endTable = System::DynamicCast<Table>(doc->get_LastSection()->GetChild(NodeType::Table, 0, true));

    // Extract the content between these nodes in the document. Include these markers in the extraction.
    std::vector<System::SharedPtr<Node>> extractedNodes = ExtractContent(startPara, endTable, true);

    // Lets reverse the array to make inserting the content back into the document easier.
    for (auto it = extractedNodes.rbegin(); it != extractedNodes.rend(); ++it)
    {
        endTable->get_ParentNode()->InsertAfter(*it, endTable);
    }

    System::String outputPath = outputDataDir + u"ExtractContentBetweenBlockLevelNodes.doc";
    // Save the generated document to disk.
    doc->Save(outputPath);
    // ExEnd:ExtractContentBetweenBlockLevelNodes
    std::cout << "Extracted content betweenn the block level nodes successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ExtractContentBetweenBlockLevelNodes example finished." << std::endl << std::endl;
}
