#include "stdafx.h"
#include "examples.h"

#include "Model/Text/Paragraph.h"
#include "Model/Sections/Section.h"
#include "Model/Sections/Body.h"
#include "Model/Nodes/NodeType.h"
#include "Model/Nodes/Node.h"
#include "Model/Document/Document.h"

using namespace Aspose::Words;

void WorkingWithNodes()
{
    std::cout << "WorkingWithNodes example started." << std::endl;
    // Create a new document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();

    // Creates and adds a paragraph node to the document.
    System::SharedPtr<Paragraph> para = System::MakeObject<Paragraph>(doc);

    // Typed access to the last section of the document.
    System::SharedPtr<Section> section = doc->get_LastSection();
    section->get_Body()->AppendChild(para);

    // Next print the node type of one of the nodes in the document.
    NodeType nodeType = doc->get_FirstSection()->get_Body()->get_NodeType();
    std::cout << "NodeType: " << Node::NodeTypeToString(nodeType).ToUtf8String() << std::endl;
    std::cout << "WorkingWithNodes example finished." << std::endl << std::endl;
}