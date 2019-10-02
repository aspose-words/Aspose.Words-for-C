#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeStart.h>

using namespace Aspose::Words;

void RemoveRegionText()
{
    std::cout << "RemoveRegionText example started." << std::endl;
    // ExStart:RemoveRegionText
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithComments();
    System::String outputDataDir = GetOutputDataDir_WorkingWithComments();

    // Open the document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");

    System::SharedPtr<CommentRangeStart> commentStart = System::DynamicCast<CommentRangeStart>(doc->GetChild(NodeType::CommentRangeStart, 0, true));
    System::SharedPtr<CommentRangeEnd> commentEnd = System::DynamicCast<CommentRangeEnd>(doc->GetChild(NodeType::CommentRangeEnd, 0, true));

    System::SharedPtr<Node> currentNode = commentStart;
    bool isRemoving = true;
    while (currentNode != nullptr && isRemoving)
    {
        if (currentNode->get_NodeType() == NodeType::CommentRangeEnd)
        {
            isRemoving = false;
        }

        System::SharedPtr<Node> nextNode = currentNode->NextPreOrder(doc);
        currentNode->Remove();
        currentNode = nextNode;
    }

    System::String outputPath = outputDataDir + u"RemoveRegionText.doc";
    // Save the document.
    doc->Save(outputPath);
    // ExEnd:RemoveRegionText
    std::cout << "Comments added successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "RemoveRegionText example finished." << std::endl << std::endl;
}