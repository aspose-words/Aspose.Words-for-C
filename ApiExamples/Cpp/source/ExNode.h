#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Nodes/INodeChangingCallback.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeChangingAction.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeChangingArgs.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Inline.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <drawing/image.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;

namespace ApiExamples {

class ExNode : public ApiExampleBase
{
public:
    void CloneCompositeNode()
    {
        //ExStart
        //ExFor:Node
        //ExFor:Node.Clone
        //ExSummary:Shows how to clone composite nodes with and without their child nodes.
        auto doc = MakeObject<Document>();
        SharedPtr<Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
        para->AppendChild(MakeObject<Run>(doc, u"Hello world!"));

        // Clone the paragraph and the child nodes
        SharedPtr<Node> cloneWithChildren = para->Clone(true);

        ASSERT_TRUE((System::DynamicCast<CompositeNode>(cloneWithChildren))->get_HasChildNodes());
        ASSERT_EQ(u"Hello world!", cloneWithChildren->GetText().Trim());

        // Clone the paragraph without its child nodes
        SharedPtr<Node> cloneWithoutChildren = para->Clone(false);

        ASSERT_FALSE((System::DynamicCast<CompositeNode>(cloneWithoutChildren))->get_HasChildNodes());
        ASSERT_EQ(String::Empty, cloneWithoutChildren->GetText().Trim());
        //ExEnd
    }

    void GetParentNode()
    {
        //ExStart
        //ExFor:Node.ParentNode
        //ExSummary:Shows how to access the parent node.
        auto doc = MakeObject<Document>();

        // Get the document's first paragraph and append a child node to it in the form of a run with text
        SharedPtr<Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();

        // When inserting a new node, the document that the node will belong to must be provided as an argument
        auto run = MakeObject<Run>(doc, u"Hello world!");
        para->AppendChild(run);

        // The node lineage can be traced back to the document itself
        ASPOSE_ASSERT_EQ(para, run->get_ParentNode());
        ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body(), para->get_ParentNode());
        ASPOSE_ASSERT_EQ(doc->get_FirstSection(), doc->get_FirstSection()->get_Body()->get_ParentNode());
        ASPOSE_ASSERT_EQ(doc, doc->get_FirstSection()->get_ParentNode());
        //ExEnd
    }

    void OwnerDocument()
    {
        //ExStart
        //ExFor:Node.Document
        //ExFor:Node.ParentNode
        //ExSummary:Shows how to create a node and set its owning document.
        // Open a file from disk
        auto doc = MakeObject<Document>();

        // Creating a new node of any type requires a document passed into the constructor
        auto para = MakeObject<Paragraph>(doc);

        // The new paragraph node does not yet have a parent
        std::cout << (String(u"Paragraph has no parent node: ") + (para->get_ParentNode() == nullptr)) << std::endl;

        // But the paragraph node knows its document
        std::cout << (String(u"Both nodes' documents are the same: ") + (para->get_Document() == doc)) << std::endl;

        // The fact that a node always belongs to a document allows us to access and modify
        // properties that reference the document-wide data such as styles or lists
        para->get_ParagraphFormat()->set_StyleName(u"Heading 1");

        // Now add the paragraph to the main text of the first section
        doc->get_FirstSection()->get_Body()->AppendChild(para);

        // The paragraph node is now a child of the Body node
        std::cout << (String(u"Paragraph has a parent node: ") + (para->get_ParentNode() != nullptr)) << std::endl;
        //ExEnd

        ASPOSE_ASSERT_EQ(doc, para->get_Document());
        ASSERT_FALSE(para->get_ParentNode() == nullptr);
    }

    void EnumerateChildNodes()
    {
        //ExStart
        //ExFor:Node
        //ExFor:NodeType
        //ExFor:CompositeNode
        //ExFor:CompositeNode.GetChild
        //ExFor:CompositeNode.ChildNodes
        //ExFor:CompositeNode.GetEnumerator
        //ExSummary:Shows how to enumerate immediate children of a CompositeNode using the enumerator provided by the ChildNodes collection.
        auto doc = MakeObject<Document>();

        auto paragraph = System::DynamicCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));
        paragraph->AppendChild(MakeObject<Run>(doc, u"Hello world!"));
        paragraph->AppendChild(MakeObject<Run>(doc, u" Hello again!"));

        SharedPtr<NodeCollection> children = paragraph->get_ChildNodes();

        // Paragraph may contain children of various types such as runs, shapes and so on
        for (auto child : System::IterateOver(children))
        {
            if (child->get_NodeType() == NodeType::Run)
            {
                auto run = System::DynamicCast<Run>(child);
                std::cout << run->get_Text() << std::endl;
            }
        }
        //ExEnd

        ASSERT_EQ(NodeType::Run, paragraph->GetChild(NodeType::Run, 0, true)->get_NodeType());
        ASSERT_EQ(2, paragraph->get_ChildNodes()->get_Count());
        ASSERT_EQ(u"Hello world! Hello again!", doc->GetText().Trim());
    }

    void IndexChildNodes()
    {
        //ExStart
        //ExFor:NodeCollection.Count
        //ExFor:NodeCollection.Item
        //ExSummary:Shows how to enumerate immediate children of a CompositeNode using indexed access.
        auto doc = MakeObject<Document>();
        auto paragraph = System::DynamicCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));
        paragraph->AppendChild(MakeObject<Run>(doc, u"Hello world!"));

        SharedPtr<NodeCollection> children = paragraph->get_ChildNodes();

        for (int i = 0; i < children->get_Count(); i++)
        {
            SharedPtr<Node> child = children->idx_get(i);

            // Paragraph may contain children of various types such as runs, shapes and so on
            if (child->get_NodeType() == NodeType::Run)
            {
                auto run = System::DynamicCast<Run>(child);
                std::cout << run->get_Text() << std::endl;
            }
        }
        //ExEnd

        ASSERT_EQ(1, paragraph->get_ChildNodes()->get_Count());
    }

    //ExStart
    //ExFor:Node.NextSibling
    //ExFor:CompositeNode.FirstChild
    //ExFor:Node.IsComposite
    //ExFor:CompositeNode.IsComposite
    //ExFor:Node.NodeTypeToString
    //ExFor:Paragraph.NodeType
    //ExFor:Table.NodeType
    //ExFor:Node.NodeType
    //ExFor:Footnote.NodeType
    //ExFor:FormField.NodeType
    //ExFor:SmartTag.NodeType
    //ExFor:Cell.NodeType
    //ExFor:Row.NodeType
    //ExFor:Document.NodeType
    //ExFor:Comment.NodeType
    //ExFor:Run.NodeType
    //ExFor:Section.NodeType
    //ExFor:SpecialChar.NodeType
    //ExFor:Shape.NodeType
    //ExFor:FieldEnd.NodeType
    //ExFor:FieldSeparator.NodeType
    //ExFor:FieldStart.NodeType
    //ExFor:BookmarkStart.NodeType
    //ExFor:CommentRangeEnd.NodeType
    //ExFor:BuildingBlock.NodeType
    //ExFor:GlossaryDocument.NodeType
    //ExFor:BookmarkEnd.NodeType
    //ExFor:GroupShape.NodeType
    //ExFor:CommentRangeStart.NodeType
    //ExSummary:Shows how to efficiently visit all direct and indirect children of a composite node.
    void RecurseAllNodes()
    {
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // Any node that can contain child nodes, such as the document itself, is composite
        ASSERT_TRUE(doc->get_IsComposite());

        // Invoke the recursive function that will go through and print all the child nodes of a composite node
        TraverseAllNodes(doc, 0);
    }

    /// <summary>
    /// Recursively traverses a node tree while printing the type of each node with an indent depending on depth as well as the contents of all inline nodes.
    /// </summary>
    void TraverseAllNodes(SharedPtr<CompositeNode> parentNode, int depth)
    {
        // Loop through immediate children of a node
        for (SharedPtr<Node> childNode = parentNode->get_FirstChild(); childNode != nullptr; childNode = childNode->get_NextSibling())
        {
            std::cout << (String(u'\t', depth)) << Node::NodeTypeToString(childNode->get_NodeType());

            // Recurse into the node if it is a composite node
            if (childNode->get_IsComposite())
            {
                std::cout << std::endl;
                TraverseAllNodes(System::DynamicCast<CompositeNode>(childNode), depth + 1);
            }
            else if (System::ObjectExt::Is<Inline>(childNode))
            {
                std::cout << " - \"" << childNode->GetText().Trim() << "\"" << std::endl;
            }
            else
            {
                std::cout << std::endl;
            }
        }
    }
    //ExEnd

    void RemoveNodes()
    {

        //ExStart
        //ExFor:Node
        //ExFor:Node.NodeType
        //ExFor:Node.Remove
        //ExSummary:Shows how to remove all nodes of a specific type from a composite node.
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        ASSERT_EQ(2, doc->GetChildNodes(NodeType::Table, true)->get_Count());

        // Select the first child node in the body
        SharedPtr<Node> curNode = doc->get_FirstSection()->get_Body()->get_FirstChild();

        while (curNode != nullptr)
        {
            // Save the next sibling node as a variable in case we want to move to it after deleting this node
            SharedPtr<Node> nextNode = curNode->get_NextSibling();

            // A section body can contain Paragraph and Table nodes
            // If the node is a Table, remove it from the parent
            if (curNode->get_NodeType() == NodeType::Table)
            {
                curNode->Remove();
            }

            // Continue going through child nodes until null (no more siblings) is reached
            curNode = nextNode;
        }

        ASSERT_EQ(0, doc->GetChildNodes(NodeType::Table, true)->get_Count());
        //ExEnd
    }

    void EnumNextSibling()
    {

        //ExStart
        //ExFor:CompositeNode.FirstChild
        //ExFor:Node.NextSibling
        //ExFor:Node.NodeTypeToString
        //ExFor:Node.NodeType
        //ExSummary:Shows how to enumerate immediate child nodes of a composite node using NextSibling.
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // Loop starting from the first child until we reach null
        for (SharedPtr<Node> node = doc->get_FirstSection()->get_Body()->get_FirstChild(); node != nullptr; node = node->get_NextSibling())
        {
            // Output the types of the nodes that we come across
            std::cout << Node::NodeTypeToString(node->get_NodeType()) << std::endl;
        }
        //ExEnd
    }

    void TypedAccess()
    {

        //ExStart
        //ExFor:Story.Tables
        //ExFor:Table.FirstRow
        //ExFor:Table.LastRow
        //ExFor:TableCollection
        //ExSummary:Shows how to use typed properties to access nodes of the document tree.
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        // Quick typed access to all Table child nodes contained in the Body
        SharedPtr<TableCollection> tables = doc->get_FirstSection()->get_Body()->get_Tables();

        ASSERT_EQ(5, tables->idx_get(0)->get_Rows()->get_Count());
        ASSERT_EQ(4, tables->idx_get(1)->get_Rows()->get_Count());

        for (auto table : System::IterateOver(tables->LINQ_OfType<SharedPtr<Table>>()))
        {
            // Quick typed access to the first row of the table
            if (table->get_FirstRow() != nullptr)
            {
                table->get_FirstRow()->Remove();
            }

            // Quick typed access to the last row of the table
            if (table->get_LastRow() != nullptr)
            {
                table->get_LastRow()->Remove();
            }
        }

        // Each table has shrunk by two rows
        ASSERT_EQ(3, tables->idx_get(0)->get_Rows()->get_Count());
        ASSERT_EQ(2, tables->idx_get(1)->get_Rows()->get_Count());
        //ExEnd
    }

    void RemoveChild()
    {
        //ExStart
        //ExFor:CompositeNode.LastChild
        //ExFor:Node.PreviousSibling
        //ExFor:CompositeNode.RemoveChild
        //ExSummary:Shows how to use of methods of Node and CompositeNode to remove a section before the last section in the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a second section by inserting a section break and add text to both sections
        builder->Writeln(u"Section 1 text.");
        builder->InsertBreak(BreakType::SectionBreakContinuous);
        builder->Writeln(u"Section 2 text.");

        // Both sections are siblings of each other
        auto lastSection = System::DynamicCast<Section>(doc->get_LastChild());
        auto firstSection = System::DynamicCast<Section>(lastSection->get_PreviousSibling());

        // Remove a section based on its sibling relationship with another section
        if (lastSection->get_PreviousSibling() != nullptr)
        {
            doc->RemoveChild(firstSection);
        }

        // The section we removed was the first one, leaving the document with only the second
        ASSERT_EQ(u"Section 2 text.", doc->GetText().Trim());
        //ExEnd
    }

    void CreateAndAddParagraphNode()
    {
        auto doc = MakeObject<Document>();

        auto para = MakeObject<Paragraph>(doc);

        SharedPtr<Section> section = doc->get_LastSection();
        section->get_Body()->AppendChild(para);
    }

    void RemoveSmartTagsFromCompositeNode()
    {
        //ExStart
        //ExFor:CompositeNode.RemoveSmartTags
        //ExSummary:Removes all smart tags from descendant nodes of the composite node.
        auto doc = MakeObject<Document>(MyDir + u"Smart tags.doc");
        ASSERT_EQ(8, doc->GetChildNodes(NodeType::SmartTag, true)->get_Count());

        // Remove smart tags from the whole document
        doc->RemoveSmartTags();

        ASSERT_EQ(0, doc->GetChildNodes(NodeType::SmartTag, true)->get_Count());
        //ExEnd
    }

    void GetIndexOfNode()
    {
        //ExStart
        //ExFor:CompositeNode.IndexOf
        //ExSummary:Shows how to get the index of a given child node from its parent.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Get the body of the first section in the document
        SharedPtr<Body> body = doc->get_FirstSection()->get_Body();

        // Retrieve the index of the last paragraph in the body
        ASSERT_EQ(24, body->get_ChildNodes()->IndexOf(body->get_LastParagraph()));
        //ExEnd
    }

    void ConvertNodeToHtmlWithDefaultOptions()
    {
        //ExStart
        //ExFor:Node.ToString(SaveFormat)
        //ExFor:Node.ToString(SaveOptions)
        //ExSummary:Exports the content of a node to String in HTML format.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Extract the last paragraph in the document to convert to HTML
        SharedPtr<Node> node = doc->get_LastSection()->get_Body()->get_LastParagraph();

        // When ToString is called using the html SaveFormat overload then the node is converted directly to html
        ASSERT_EQ(String(u"<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%; font-size:12pt\">") +
                      u"<span style=\"font-family:'Times New Roman'\">Hello World!</span>" + u"</p>",
                  node->ToString(SaveFormat::Html));

        // We can also modify the result of this conversion using a SaveOptions object
        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_ExportRelativeFontSize(true);

        ASSERT_EQ(String(u"<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%\">") +
                      u"<span style=\"font-family:'Times New Roman'\">Hello World!</span>" + u"</p>",
                  node->ToString(saveOptions));
        //ExEnd
    }

    void TypedNodeCollectionToArray()
    {
        //ExStart
        //ExFor:ParagraphCollection.ToArray
        //ExSummary:Shows how to create an array from a NodeCollection.
        // You can use ToArray to return a typed array of nodes
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        ArrayPtr<SharedPtr<Paragraph>> paras = doc->get_FirstSection()->get_Body()->get_Paragraphs()->ToArray();

        ASSERT_EQ(22, paras->get_Length());
        //ExEnd
    }

    void NodeEnumerationHotRemove()
    {
        //ExStart
        //ExFor:ParagraphCollection.ToArray
        //ExSummary:Shows how to use "hot remove" to remove a node during enumeration.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"The first paragraph");
        builder->Writeln(u"The second paragraph");
        builder->Writeln(u"The third paragraph");
        builder->Writeln(u"The fourth paragraph");

        // Hot remove allows a node to be removed from a live collection and have the enumeration continue
        for (SharedPtr<Paragraph> para : doc->get_FirstSection()->get_Body()->get_Paragraphs()->ToArray())
        {
            if (para->get_Range()->get_Text().Contains(u"third"))
            {
                para->Remove();
            }
        }

        ASSERT_FALSE(doc->GetText().Contains(u"The third paragraph"));
        //ExEnd
    }

protected:
    void TestNodeXPathNavigator(String navigatorResult, SharedPtr<Document> doc)
    {
        for (auto run : System::IterateOver(doc->GetChildNodes(NodeType::Run, true)->ToArray()->LINQ_OfType<SharedPtr<Run>>()))
        {
            ASSERT_TRUE(navigatorResult.Contains(run->GetText().Trim()));
        }
    }

public:
    //ExStart
    //ExFor:NodeChangingAction
    //ExFor:NodeChangingArgs.Action
    //ExFor:NodeChangingArgs.NewParent
    //ExFor:NodeChangingArgs.OldParent
    //ExSummary:Shows how to use a NodeChangingCallback to monitor changes to the document tree as it is edited.
    void NodeChangingCallback()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set the NodeChangingCallback attribute to a custom printer
        doc->set_NodeChangingCallback(MakeObject<ExNode::NodeChangingPrinter>());

        // All node additions and removals will be printed to the console as we edit the document
        builder->Writeln(u"Hello world!");
        builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Cell 1");
        builder->InsertCell();
        builder->Write(u"Cell 2");
        builder->EndTable();

        builder->InsertImage(System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg"));

        builder->get_CurrentParagraph()->get_ParentNode()->RemoveAllChildren();
    }

private:
    /// <summary>
    /// Prints every node insertion/removal as it takes place in the document.
    /// </summary>
    class NodeChangingPrinter : public INodeChangingCallback
    {
    public:
        void NodeInserting(SharedPtr<NodeChangingArgs> args) override
        {
            ASSERT_EQ(NodeChangingAction::Insert, args->get_Action());
            ASPOSE_ASSERT_EQ(nullptr, args->get_OldParent());
        }

        void NodeInserted(SharedPtr<NodeChangingArgs> args) override
        {
            ASSERT_EQ(NodeChangingAction::Insert, args->get_Action());
            ASSERT_FALSE(args->get_NewParent() == nullptr);

            std::cout << "Inserted node:" << std::endl;
            std::cout << String::Format(u"\tType:\t{0}", args->get_Node()->get_NodeType()) << std::endl;

            if (args->get_Node()->GetText().Trim() != u"")
            {
                std::cout << "\tText:\t\"" << args->get_Node()->GetText().Trim() << "\"" << std::endl;
            }

            std::cout << "\tHash:\t" << System::ObjectExt::GetHashCode(args->get_Node()) << std::endl;
            std::cout << String::Format(u"\tParent:\t{0} ({1})", args->get_NewParent()->get_NodeType(),
                                                System::ObjectExt::GetHashCode(args->get_NewParent()))
                      << std::endl;
        }

        void NodeRemoving(SharedPtr<NodeChangingArgs> args) override
        {
            ASSERT_EQ(NodeChangingAction::Remove, args->get_Action());
        }

        void NodeRemoved(SharedPtr<NodeChangingArgs> args) override
        {
            ASSERT_EQ(NodeChangingAction::Remove, args->get_Action());
            ASSERT_TRUE(args->get_NewParent() == nullptr);

            std::cout << String::Format(u"Removed node: {0} ({1})", args->get_Node()->get_NodeType(), System::ObjectExt::GetHashCode(args->get_Node()))
                      << std::endl;
        }
    };
    //ExEnd

public:
    void NodeCollection_()
    {
        //ExStart
        //ExFor:NodeCollection.Contains(Node)
        //ExFor:NodeCollection.Insert(Int32,Node)
        //ExFor:NodeCollection.Remove(Node)
        //ExSummary:Shows how to work with a NodeCollection.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // The normal way to insert Runs into a document is to add text using a DocumentBuilder
        builder->Write(u"Run 1. ");
        builder->Write(u"Run 2. ");

        // Every .Write() invocation creates a new Run, which is added to the parent Paragraph's RunCollection
        SharedPtr<RunCollection> runs = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs();
        ASSERT_EQ(2, runs->get_Count());

        // We can insert a node into the RunCollection manually to achieve the same effect
        auto newRun = MakeObject<Run>(doc, u"Run 3. ");
        runs->Insert(3, newRun);

        ASSERT_TRUE(runs->Contains(newRun));
        ASSERT_EQ(u"Run 1. Run 2. Run 3.", doc->GetText().Trim());

        // Text can also be deleted from the document by accessing individual Runs via the RunCollection and editing or removing them
        SharedPtr<Run> run = runs->idx_get(1);
        runs->Remove(run);
        ASSERT_EQ(u"Run 1. Run 3.", doc->GetText().Trim());

        ASSERT_FALSE(run == nullptr);
        ASSERT_FALSE(runs->Contains(run));
        //ExEnd
    }

};

} // namespace ApiExamples
