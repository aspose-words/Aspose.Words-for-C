#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace
{
    void UseNodeType()
    {
        // ExStart:UseNodeType
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        // Returns NodeType.Document
        NodeType type = doc->get_NodeType();
        // ExEnd:UseNodeType
    }

    void GetParentNode()
    {
        // ExStart:GetParentNode
        // Create a new empty document. It has one section.
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        // The section is the first child node of the document.
        System::SharedPtr<Node> section = doc->get_FirstChild();
        // The section's parent node is the document.
        std::cout << "Section parent is the document: " << System::ObjectExt::Box<bool>((doc == section->get_ParentNode()))->ToString().ToUtf8String() << std::endl;
        // ExEnd:GetParentNode
    }

    void OwnerDocument()
    {
        // ExStart:OwnerDocument
        // Open a file from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        // Creating a new node of any type requires a document passed into the constructor.
        System::SharedPtr<Paragraph> para = System::MakeObject<Paragraph>(doc);
        // The new paragraph node does not yet have a parent.
        std::cout << "Paragraph has no parent node: " << System::ObjectExt::Box<bool>((para->get_ParentNode() == nullptr))->ToString().ToUtf8String() << std::endl;
        // But the paragraph node knows its document.
        std::cout << "Both nodes' documents are the same: " << System::ObjectExt::Box<bool>((para->get_Document() == doc))->ToString().ToUtf8String() << std::endl;
        // The fact that a node always belongs to a document allows us to access and modify
        // Properties that reference the document-wide data such as styles or lists.
        para->get_ParagraphFormat()->set_StyleName(u"Heading 1");
        // Now add the paragraph to the main text of the first section.
        doc->get_FirstSection()->get_Body()->AppendChild(para);
        // The paragraph node is now a child of the Body node.
        std::cout << "Paragraph has a parent node: " << System::ObjectExt::Box<bool>((para->get_ParentNode() != nullptr))->ToString().ToUtf8String() << std::endl;
        // ExEnd:OwnerDocument
    }

    void EnumerateChildNodes()
    {
        // ExStart:EnumerateChildNodes
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<Paragraph> paragraph = System::DynamicCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));
        System::SharedPtr<NodeCollection> children = paragraph->get_ChildNodes();
        for (System::SharedPtr<Node> child : System::IterateOver(children))
        {
            if (System::ObjectExt::Equals(child->get_NodeType(), NodeType::Run))
            {
                // Say we found the node that we want, do something useful.
                System::SharedPtr<Run> run = System::DynamicCast<Run>(child);
                std::cout << run->get_Text().ToUtf8String() << std::endl;
            }
        }
        // ExEnd:EnumerateChildNodes
    }

    void IndexChildNodes()
    {
        // ExStart:IndexChildNodes
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<Paragraph> paragraph = System::DynamicCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));
        System::SharedPtr<NodeCollection> children = paragraph->get_ChildNodes();
        for (int32_t i = 0; i < children->get_Count(); i++)
        {
            System::SharedPtr<Node> child = children->idx_get(i);
            // Paragraph may contain children of various types such as runs, shapes and so on.
            if (System::ObjectExt::Equals(child->get_NodeType(), NodeType::Run))
            {
                // Say we found the node that we want, do something useful.
                System::SharedPtr<Run> run = System::DynamicCast<Run>(child);
                std::cout << run->get_Text().ToUtf8String() << std::endl;
            }
        }
        // ExEnd:IndexChildNodes
    }

    // ExStart:RecurseAllNodes
    void TraverseAllNodes(System::SharedPtr<CompositeNode> parentNode)
    {
        // This is the most efficient way to loop through immediate children of a node.
        for (System::SharedPtr<Node> childNode = parentNode->get_FirstChild(); childNode != nullptr; childNode = childNode->get_NextSibling())
        {
            // Do some useful work.
            std::cout << Node::NodeTypeToString(childNode->get_NodeType()).ToUtf8String() << std::endl;
            // Recurse into the node if it is a composite node.
            if (childNode->get_IsComposite())
            {
                TraverseAllNodes(System::DynamicCast<CompositeNode>(childNode));
            }
        }
    }

    void RecurseAllNodes(System::String const &inputDataDir)
    {
        // Open a document.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Node.RecurseAllNodes.doc");
        // Invoke the recursive function that will walk the tree.
        TraverseAllNodes(doc);
    }
    // ExEnd:RecurseAllNodes

    void TypedAccess()
    {
        // ExStart:TypedAccess
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<Section> section = doc->get_FirstSection();
        // Quick typed access to the Body child node of the Section.
        System::SharedPtr<Body> body = section->get_Body();
        // Quick typed access to all Table child nodes contained in the Body.
        System::SharedPtr<TableCollection> tables = body->get_Tables();
        for (System::SharedPtr<Table> table : System::IterateOver<System::SharedPtr<Table>>(tables))
        {
            // Quick typed access to the first row of the table.
            if (table->get_FirstRow() != nullptr)
            {
                table->get_FirstRow()->Remove();
            }
            // Quick typed access to the last row of the table.
            if (table->get_LastRow() != nullptr)
            {
                table->get_LastRow()->Remove();
            }
        }
        // ExEnd:TypedAccess
    }

    void CreateAndAddParagraphNode()
    {
        // ExStart:CreateAndAddParagraphNode
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<Paragraph> para = System::MakeObject<Paragraph>(doc);
        System::SharedPtr<Section> section = doc->get_LastSection();
        section->get_Body()->AppendChild(para);
        // ExEnd:CreateAndAddParagraphNode
    }
}

void ExNode()
{
    std::cout << "ExNode example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithNode();
    // The following method shows how to use the NodeType enumeration.
    UseNodeType();
    // The following method shows how to access the parent node.
    GetParentNode();
    // The following method shows that when you create any node, it requires a document that will own the node.
    OwnerDocument();
    // Shows how to extract a specific child node from a CompositeNode by using the GetChild method and passing the NodeType and index.
    EnumerateChildNodes();
    // Shows how to enumerate immediate children of a CompositeNode using indexed access.
    IndexChildNodes();
    // Shows how to efficiently visit all direct and indirect children of a composite node.
    RecurseAllNodes(inputDataDir);
    // Demonstrates how to use typed properties to access nodes of the document tree.
    TypedAccess();
    // The following method shows how to creates and adds a paragraph node.
    CreateAndAddParagraphNode();
    std::cout << "ExNode example finished." << std::endl << std::endl;
}