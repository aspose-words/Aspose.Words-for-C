#include "stdafx.h"
#include "examples.h"

#include <Model/Text/RunCollection.h>
#include <Model/Text/Run.h>
#include <Model/Text/Paragraph.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/Node.h>
#include <Model/Document/Document.h>
#include <Model/Document/SaveFormat.h>

#include "Common.h"

using namespace Aspose::Words;

void ExtractContentBetweenRuns()
{
    std::cout << "ExtractContentBetweenRuns example started." << std::endl;
    // ExStart:ExtractContentBetweenRuns
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");

    // Retrieve a paragraph from the first section.
    System::SharedPtr<Paragraph> para = System::DynamicCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 7, true));

    // Use some runs for extraction.
    System::SharedPtr<Run> startRun = para->get_Runs()->idx_get(1);
    System::SharedPtr<Run> endRun = para->get_Runs()->idx_get(4);

    // Extract the content between these nodes in the document. Include these markers in the extraction.
    std::vector<System::SharedPtr<Node>> extractedNodes = ExtractContent(startRun, endRun, true);

    // Get the node from the list. There should only be one paragraph returned in the list.
    System::SharedPtr<Node> node = extractedNodes[0];
    // Print the text of this node to the console.
    std::cout << node->ToString(SaveFormat::Text).ToUtf8String() << std::endl;
    // ExEnd:ExtractContentBetweenRuns
    std::cout << "Extracted content between the runs successfully." << std::endl;
    std::cout << "ExtractContentBetweenRuns example finished." << std::endl << std::endl;
}
