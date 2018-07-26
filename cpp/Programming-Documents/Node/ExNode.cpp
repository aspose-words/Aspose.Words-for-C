#include "../../examples.h"

#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/special_casts.h>

#include "Model/Document/Document.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/TableCollection.h>
#include <Model/Tables/Row.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/Text/Run.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/Section.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/CompositeNode.h>

using namespace Aspose::Words;

namespace
{
    void UseNodeType()
    {
        // ExStart:UseNodeType
        auto doc = System::MakeObject<Document>();
        // Returns NodeType.Document
        NodeType type = doc->get_NodeType();
        // ExEnd:UseNodeType
    }

    void GetParentNode()
    {
        // ExStart:GetParentNode
        // Create a new empty document. It has one section.
        auto doc = System::MakeObject<Document>();
        // The section is the first child node of the document.
        auto section = doc->get_FirstChild();
        // The section's parent node is the document.
        std::cout << "Section parent is the document: " << System::ObjectExt::Box<bool>((doc == section->get_ParentNode()))->ToString().ToUtf8String() << std::endl;
        // ExEnd:GetParentNode
    }

    void OwnerDocument()
    {
        // ExStart:OwnerDocument
        // Open a file from disk.
        auto doc = System::MakeObject<Document>();
        // Creating a new node of any type requires a document passed into the constructor.
        auto para = System::MakeObject<Paragraph>(doc);
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
        auto doc = System::MakeObject<Document>();
        auto paragraph = System::DynamicCast<Aspose::Words::Paragraph>(doc->GetChild(Aspose::Words::NodeType::Paragraph, 0, true));
        auto children = paragraph->get_ChildNodes();
        auto child_enumerator = (children)->GetEnumerator();
        decltype(child_enumerator->get_Current()) child;
        while (child_enumerator->MoveNext() && (child = child_enumerator->get_Current(), true))
        {
            // Paragraph may contain children of various types such as runs, shapes and so on.
            if (System::ObjectExt::Equals(child->get_NodeType(), Aspose::Words::NodeType::Run))
            {
                // Say we found the node that we want, do something useful.
                auto run = System::DynamicCast<Aspose::Words::Run>(child);
                std::cout << run->get_Text().ToUtf8String() << std::endl;
            }
        }
        // ExEnd:EnumerateChildNodes
    }

    void IndexChildNodes()
    {
        // ExStart:IndexChildNodes
        auto doc = System::MakeObject<Document>();
        auto paragraph = System::DynamicCast<Aspose::Words::Paragraph>(doc->GetChild(Aspose::Words::NodeType::Paragraph, 0, true));
        auto children = paragraph->get_ChildNodes();
        for (int32_t i = 0; i < children->get_Count(); i++)
        {
            auto child = children->idx_get(i);
            // Paragraph may contain children of various types such as runs, shapes and so on.
            if (System::ObjectExt::Equals(child->get_NodeType(), Aspose::Words::NodeType::Run))
            {
                // Say we found the node that we want, do something useful.
                auto run = System::DynamicCast<Aspose::Words::Run>(child);
                std::cout << run->get_Text().ToUtf8String() << std::endl;
            }
        }
        // ExEnd:IndexChildNodes
    }    

	// ExStart:RecurseAllNodes
    void RecurseAllNodes()
    {
        // The path to the documents directory.
        System::String dataDir = GetDataDir_WorkingWithNode();
        // Open a document.
        auto doc = System::MakeObject<Document>(dataDir + u"Node.RecurseAllNodes.doc");
        // Invoke the recursive function that will walk the tree.
        TraverseAllNodes(doc);
    }

	void TraverseAllNodes(System::SharedPtr<CompositeNode> parentNode)
	{
		// This is the most efficient way to loop through immediate children of a node.
		for (auto childNode = parentNode->get_FirstChild(); childNode != nullptr; childNode = childNode->get_NextSibling())
		{
			// Do some useful work.
			std::cout << Node::NodeTypeToString(childNode->get_NodeType()).ToUtf8String() << std::endl;
			// Recurse into the node if it is a composite node.
			if (childNode->get_IsComposite())
			{
				TraverseAllNodes(System::DynamicCast<Aspose::Words::CompositeNode>(childNode));
			}
		}
	}
	// ExEnd:RecurseAllNodes

    void TypedAccess()
    {
        // ExStart:TypedAccess
        auto doc = System::MakeObject<Document>();
        auto section = doc->get_FirstSection();
        // Quick typed access to the Body child node of the Section.
        auto body = section->get_Body();
        // Quick typed access to all Table child nodes contained in the Body.
        auto tables = body->get_Tables();
        auto table_enumerator = (System::DynamicCastEnumerableTo<System::SharedPtr<Aspose::Words::Tables::Table>>(tables))->GetEnumerator();
        decltype(table_enumerator->get_Current()) table;
        while (table_enumerator->MoveNext() && (table = table_enumerator->get_Current(), true))
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
        auto doc = System::MakeObject<Document>();
        auto para = System::MakeObject<Paragraph>(doc);
        auto section = doc->get_LastSection();
        section->get_Body()->AppendChild(para);
        // ExEnd:CreateAndAddParagraphNode
    }
}

void ExNode()
{
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
    RecurseAllNodes();
    // Demonstrates how to use typed properties to access nodes of the document tree.
    TypedAccess();
    // The following method shows how to creates and adds a paragraph node.
    CreateAndAddParagraphNode();
}