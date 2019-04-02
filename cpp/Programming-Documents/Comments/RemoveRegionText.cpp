#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Text/CommentRangeEnd.h>
#include <Model/Text/CommentRangeStart.h>

using namespace Aspose::Words;

void RemoveRegionText()
{
    std::cout << "RemoveRegionText example started." << std::endl;
    // ExStart:RemoveRegionText
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithComments();

    // Open the document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");

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

    System::String outputPath = dataDir + GetOutputFilePath(u"RemoveRegionText.doc");
    // Save the document.
    doc->Save(outputPath);
    // ExEnd:RemoveRegionText
    std::cout << "Comments added successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "RemoveRegionText example finished." << std::endl << std::endl;
}