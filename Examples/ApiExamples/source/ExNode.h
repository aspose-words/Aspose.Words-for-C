#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/INodeChangingCallback.h>
#include <Aspose.Words.Cpp/Inline.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeChangingAction.h>
#include <Aspose.Words.Cpp/NodeChangingArgs.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/Tables/TableCollection.h>
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

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

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
        //ExSummary:Shows how to clone a composite node.
        auto doc = MakeObject<Document>();
        SharedPtr<Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
        para->AppendChild(MakeObject<Run>(doc, u"Hello world!"));

        // Below are two ways of cloning a composite node.
        // 1 -  Create a clone of a node, and create a clone of each of its child nodes as well.
        SharedPtr<Node> cloneWithChildren = para->Clone(true);

        ASSERT_TRUE((System::ExplicitCast<CompositeNode>(cloneWithChildren))->get_HasChildNodes());
        ASSERT_EQ(u"Hello world!", cloneWithChildren->GetText().Trim());

        // 2 -  Create a clone of a node just by itself without any children.
        SharedPtr<Node> cloneWithoutChildren = para->Clone(false);

        ASSERT_FALSE((System::ExplicitCast<CompositeNode>(cloneWithoutChildren))->get_HasChildNodes());
        ASSERT_EQ(String::Empty, cloneWithoutChildren->GetText().Trim());
        //ExEnd
    }

    void GetParentNode()
    {
        //ExStart
        //ExFor:Node.ParentNode
        //ExSummary:Shows how to access a node's parent node.
        auto doc = MakeObject<Document>();
        SharedPtr<Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();

        // Append a child Run node to the document's first paragraph.
        auto run = MakeObject<Run>(doc, u"Hello world!");
        para->AppendChild(run);

        // The paragraph is the parent node of the run node. We can trace this lineage
        // all the way to the document node, which is the root of the document's node tree.
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
        auto doc = MakeObject<Document>();
        auto para = MakeObject<Paragraph>(doc);
        para->AppendChild(MakeObject<Run>(doc, u"Hello world!"));

        // We have not yet appended this paragraph as a child to any composite node.
        ASSERT_TRUE(para->get_ParentNode() == nullptr);

        // If a node is an appropriate child node type of another composite node,
        // we can attach it as a child only if both nodes have the same owner document.
        // The owner document is the document we passed to the node's constructor.
        // We have not attached this paragraph to the document, so the document does not contain its text.
        ASPOSE_ASSERT_EQ(para->get_Document(), doc);
        ASSERT_EQ(String::Empty, doc->GetText().Trim());

        // Since the document owns this paragraph, we can apply one of its styles to the paragraph's contents.
        para->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));

        // Add this node to the document, and then verify its contents.
        doc->get_FirstSection()->get_Body()->AppendChild(para);

        ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body(), para->get_ParentNode());
        ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
        //ExEnd

        ASPOSE_ASSERT_EQ(doc, para->get_Document());
        ASSERT_FALSE(para->get_ParentNode() == nullptr);
    }

    void ChildNodesEnumerate()
    {
        //ExStart
        //ExFor:Node
        //ExFor:Node.CustomNodeId
        //ExFor:NodeType
        //ExFor:CompositeNode
        //ExFor:CompositeNode.GetChild
        //ExFor:CompositeNode.ChildNodes
        //ExFor:CompositeNode.GetEnumerator
        //ExFor:NodeCollection.Count
        //ExFor:NodeCollection.Item
        //ExSummary:Shows how to traverse through a composite node's collection of child nodes.
        auto doc = MakeObject<Document>();

        // Add two runs and one shape as child nodes to the first paragraph of this document.
        auto paragraph = System::ExplicitCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));
        paragraph->AppendChild(MakeObject<Run>(doc, u"Hello world! "));

        auto shape = MakeObject<Shape>(doc, ShapeType::Rectangle);
        shape->set_Width(200);
        shape->set_Height(200);
        // Note that the 'CustomNodeId' is not saved to an output file and exists only during the node lifetime.
        shape->set_CustomNodeId(100);
        shape->set_WrapType(WrapType::Inline);
        paragraph->AppendChild(shape);

        paragraph->AppendChild(MakeObject<Run>(doc, u"Hello again!"));

        // Iterate through the paragraph's collection of immediate children,
        // and print any runs or shapes that we find within.
        SharedPtr<NodeCollection> children = paragraph->get_ChildNodes();

        ASSERT_EQ(3, paragraph->get_ChildNodes()->get_Count());

        for (const auto& child : System::IterateOver(children))
        {
            switch (child->get_NodeType())
            {
            case NodeType::Run:
                std::cout << "Run contents:" << std::endl;
                std::cout << "\t\"" << child->GetText().Trim() << "\"" << std::endl;
                break;

            case NodeType::Shape: {
                auto childShape = System::ExplicitCast<Shape>(child);
                std::cout << "Shape:" << std::endl;
                std::cout << String::Format(u"\t{0}, {1}x{2}", childShape->get_ShapeType(), childShape->get_Width(), childShape->get_Height()) << std::endl;
                ASSERT_EQ(100, shape->get_CustomNodeId());
                break;
            }

            default:
                break;
            }
        }
        //ExEnd

        ASSERT_EQ(NodeType::Run, paragraph->GetChild(NodeType::Run, 0, true)->get_NodeType());
        ASSERT_EQ(u"Hello world! Hello again!", doc->GetText().Trim());
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
    //ExSummary:Shows how to traverse a composite node's tree of child nodes.
    void RecurseChildren()
    {
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // Any node that can contain child nodes, such as the document itself, is composite.
        ASSERT_TRUE(doc->get_IsComposite());

        // Invoke the recursive function that will go through and print all the child nodes of a composite node.
        TraverseAllNodes(doc, 0);
    }

    /// <summary>
    /// Recursively traverses a node tree while printing the type of each node
    /// with an indent depending on depth as well as the contents of all inline nodes.
    /// </summary>
    void TraverseAllNodes(SharedPtr<CompositeNode> parentNode, int depth)
    {
        for (SharedPtr<Node> childNode = parentNode->get_FirstChild(); childNode != nullptr; childNode = childNode->get_NextSibling())
        {
            std::cout << (String(u'\t', depth)) << Node::NodeTypeToString(childNode->get_NodeType());

            // Recurse into the node if it is a composite node. Otherwise, print its contents if it is an inline node.
            if (childNode->get_IsComposite())
            {
                std::cout << std::endl;
                TraverseAllNodes(System::ExplicitCast<CompositeNode>(childNode), depth + 1);
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
        //ExSummary:Shows how to remove all child nodes of a specific type from a composite node.
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        ASSERT_EQ(2, doc->GetChildNodes(NodeType::Table, true)->get_Count());

        SharedPtr<Node> curNode = doc->get_FirstSection()->get_Body()->get_FirstChild();

        while (curNode != nullptr)
        {
            // Save the next sibling node as a variable in case we want to move to it after deleting this node.
            SharedPtr<Node> nextNode = curNode->get_NextSibling();

            // A section body can contain Paragraph and Table nodes.
            // If the node is a Table, remove it from the parent.
            if (curNode->get_NodeType() == NodeType::Table)
            {
                curNode->Remove();
            }

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
        //ExSummary:Shows how to use a node's NextSibling property to enumerate through its immediate children.
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        for (SharedPtr<Node> node = doc->get_FirstSection()->get_Body()->get_FirstChild(); node != nullptr; node = node->get_NextSibling())
        {
            std::cout << std::endl;
            std::cout << "Node type: " << Node::NodeTypeToString(node->get_NodeType()) << std::endl;

            String contents = node->GetText().Trim();
            std::cout << (contents == String::Empty ? u"This node contains no text" : String::Format(u"Contents: \"{0}\"", node->GetText().Trim()))
                      << std::endl;
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
        //ExSummary:Shows how to remove the first and last rows of all tables in a document.
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        SharedPtr<TableCollection> tables = doc->get_FirstSection()->get_Body()->get_Tables();

        ASSERT_EQ(5, tables->idx_get(0)->get_Rows()->get_Count());
        ASSERT_EQ(4, tables->idx_get(1)->get_Rows()->get_Count());

        for (const auto& table : System::IterateOver(tables->LINQ_OfType<SharedPtr<Table>>()))
        {
            if (table->get_FirstRow() != nullptr)
            {
                table->get_FirstRow()->Remove();
            }
            if (table->get_LastRow() != nullptr)
            {
                table->get_LastRow()->Remove();
            }
        }

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

        builder->Writeln(u"Section 1 text.");
        builder->InsertBreak(BreakType::SectionBreakContinuous);
        builder->Writeln(u"Section 2 text.");

        // Both sections are siblings of each other.
        auto lastSection = System::ExplicitCast<Section>(doc->get_LastChild());
        auto firstSection = System::ExplicitCast<Section>(lastSection->get_PreviousSibling());

        // Remove a section based on its sibling relationship with another section.
        if (lastSection->get_PreviousSibling() != nullptr)
        {
            doc->RemoveChild(firstSection);
        }

        // The section we removed was the first one, leaving the document with only the second.
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
        //ExSummary:Removes all smart tags from descendant nodes of a composite node.
        auto doc = MakeObject<Document>(MyDir + u"Smart tags.doc");

        ASSERT_EQ(8, doc->GetChildNodes(NodeType::SmartTag, true)->get_Count());

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

        SharedPtr<Body> body = doc->get_FirstSection()->get_Body();

        // Retrieve the index of the last paragraph in the body of the first section.
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

        SharedPtr<Node> node = doc->get_LastSection()->get_Body()->get_LastParagraph();

        // When we call the ToString method using the html SaveFormat overload,
        // it converts the node's contents to their raw html representation.
        ASSERT_EQ(String(u"<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%; font-size:12pt\">") +
                      u"<span style=\"font-family:'Times New Roman'\">Hello World!</span>" + u"</p>",
                  node->ToString(SaveFormat::Html));

        // We can also modify the result of this conversion using a SaveOptions object.
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

        // Remove a node from the collection in the middle of an enumeration.
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

    //ExStart
    //ExFor:NodeChangingAction
    //ExFor:NodeChangingArgs.Action
    //ExFor:NodeChangingArgs.NewParent
    //ExFor:NodeChangingArgs.OldParent
    //ExSummary:Shows how to use a NodeChangingCallback to monitor changes to the document tree in real-time as we edit it.
    void NodeChangingCallback()
    {
        auto doc = MakeObject<Document>();
        doc->set_NodeChangingCallback(MakeObject<ExNode::NodeChangingPrinter>());

        auto builder = MakeObject<DocumentBuilder>(doc);
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

    /// <summary>
    /// Prints every node insertion/removal as it takes place in the document.
    /// </summary>
    class NodeChangingPrinter : public INodeChangingCallback
    {
    private:
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
            std::cout << String::Format(u"\tParent:\t{0} ({1})", args->get_NewParent()->get_NodeType(), System::ObjectExt::GetHashCode(args->get_NewParent()))
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

    void NodeCollection_()
    {
        //ExStart
        //ExFor:NodeCollection.Contains(Node)
        //ExFor:NodeCollection.Insert(Int32,Node)
        //ExFor:NodeCollection.Remove(Node)
        //ExSummary:Shows how to work with a NodeCollection.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add text to the document by inserting Runs using a DocumentBuilder.
        builder->Write(u"Run 1. ");
        builder->Write(u"Run 2. ");

        // Every invocation of the "Write" method creates a new Run,
        // which then appears in the parent Paragraph's RunCollection.
        SharedPtr<RunCollection> runs = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs();

        ASSERT_EQ(2, runs->get_Count());

        // We can also insert a node into the RunCollection manually.
        auto newRun = MakeObject<Run>(doc, u"Run 3. ");
        runs->Insert(3, newRun);

        ASSERT_TRUE(runs->Contains(newRun));
        ASSERT_EQ(u"Run 1. Run 2. Run 3.", doc->GetText().Trim());

        // Access individual runs and remove them to remove their text from the document.
        SharedPtr<Run> run = runs->idx_get(1);
        runs->Remove(run);

        ASSERT_EQ(u"Run 1. Run 3.", doc->GetText().Trim());
        ASSERT_FALSE(run == nullptr);
        ASSERT_FALSE(runs->Contains(run));
        //ExEnd
    }
};

} // namespace ApiExamples
