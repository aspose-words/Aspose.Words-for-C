#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/Tables/TableCollection.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithNode : public DocsExamplesBase
{
public:
    void UseNodeType()
    {
        //ExStart:UseNodeType
        auto doc = MakeObject<Document>();

        NodeType type = doc->get_NodeType();
        //ExEnd:UseNodeType
    }

    void GetParentNode()
    {
        //ExStart:GetParentNode
        auto doc = MakeObject<Document>();

        // The section is the first child node of the document.
        SharedPtr<Node> section = doc->get_FirstChild();

        // The section's parent node is the document.
        std::cout << (String(u"Section parent is the document: ") + (doc == section->get_ParentNode())) << std::endl;
        //ExEnd:GetParentNode
    }

    void OwnerDocument()
    {
        //ExStart:OwnerDocument
        auto doc = MakeObject<Document>();

        // Creating a new node of any type requires a document passed into the constructor.
        auto para = MakeObject<Paragraph>(doc);

        // The new paragraph node does not yet have a parent.
        std::cout << (String(u"Paragraph has no parent node: ") + (para->get_ParentNode() == nullptr)) << std::endl;

        // But the paragraph node knows its document.
        std::cout << (String(u"Both nodes' documents are the same: ") + (para->get_Document() == doc)) << std::endl;

        // The fact that a node always belongs to a document allows us to access and modify
        // properties that reference the document-wide data, such as styles or lists.
        para->get_ParagraphFormat()->set_StyleName(u"Heading 1");

        // Now add the paragraph to the main text of the first section.
        doc->get_FirstSection()->get_Body()->AppendChild(para);

        // The paragraph node is now a child of the Body node.
        std::cout << (String(u"Paragraph has a parent node: ") + (para->get_ParentNode() != nullptr)) << std::endl;
        //ExEnd:OwnerDocument
    }

    void EnumerateChildNodes()
    {
        //ExStart:EnumerateChildNodes
        auto doc = MakeObject<Document>();
        auto paragraph = System::DynamicCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));

        SharedPtr<NodeCollection> children = paragraph->get_ChildNodes();
        for (auto child : System::IterateOver(children))
        {
            // A paragraph may contain children of various types such as runs, shapes, and others.
            if (child->get_NodeType() == NodeType::Run)
            {
                auto run = System::DynamicCast<Aspose::Words::Run>(child);
                std::cout << run->get_Text() << std::endl;
            }
        }
        //ExEnd:EnumerateChildNodes
    }

    void RecurseAllNodes()
    {
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // Invoke the recursive function that will walk the tree.
        TraverseAllNodes(doc);
    }

    /// <summary>
    /// A simple function that will walk through all children of a specified node recursively
    /// and print the type of each node to the screen.
    /// </summary>
    void TraverseAllNodes(SharedPtr<CompositeNode> parentNode)
    {
        // This is the most efficient way to loop through immediate children of a node.
        for (SharedPtr<Node> childNode = parentNode->get_FirstChild(); childNode != nullptr; childNode = childNode->get_NextSibling())
        {
            std::cout << Node::NodeTypeToString(childNode->get_NodeType()) << std::endl;

            // Recurse into the node if it is a composite node.
            if (childNode->get_IsComposite())
            {
                TraverseAllNodes(System::DynamicCast<CompositeNode>(childNode));
            }
        }
    }

    void TypedAccess()
    {
        //ExStart:TypedAccess
        auto doc = MakeObject<Document>();

        SharedPtr<Section> section = doc->get_FirstSection();
        SharedPtr<Body> body = section->get_Body();

        // Quick typed access to all Table child nodes contained in the Body.
        SharedPtr<TableCollection> tables = body->get_Tables();

        for (auto table : System::IterateOver<Table>(tables))
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
        //ExEnd:TypedAccess
    }

    void CreateAndAddParagraphNode()
    {
        //ExStart:CreateAndAddParagraphNode
        auto doc = MakeObject<Document>();

        auto para = MakeObject<Paragraph>(doc);

        SharedPtr<Section> section = doc->get_LastSection();
        section->get_Body()->AppendChild(para);
        //ExEnd:CreateAndAddParagraphNode
    }
};

}} // namespace DocsExamples::Programming_with_Documents
