// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExNode.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/enumerator_adapter.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Inline.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeChangingAction.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(718260281u, ::Aspose::Words::ApiExamples::ExNode::NodeChangingPrinter, ThisTypeBaseTypesInfo);

void ExNode::NodeChangingPrinter::NodeInserting(System::SharedPtr<Aspose::Words::NodeChangingArgs> args)
{
    ASSERT_EQ(Aspose::Words::NodeChangingAction::Insert, args->get_Action());
    ASPOSE_ASSERT_EQ(nullptr, args->get_OldParent());
}

void ExNode::NodeChangingPrinter::NodeInserted(System::SharedPtr<Aspose::Words::NodeChangingArgs> args)
{
    ASSERT_EQ(Aspose::Words::NodeChangingAction::Insert, args->get_Action());
    ASSERT_FALSE(System::TestTools::IsNull(args->get_NewParent()));
    
    std::cout << "Inserted node:" << std::endl;
    std::cout << System::String::Format(u"\tType:\t{0}", args->get_Node()->get_NodeType()) << std::endl;
    
    if (args->get_Node()->GetText().Trim() != u"")
    {
        std::cout << System::String::Format(u"\tText:\t\"{0}\"", args->get_Node()->GetText().Trim()) << std::endl;
    }
    
    std::cout << System::String::Format(u"\tHash:\t{0}", System::ObjectExt::GetHashCode(args->get_Node())) << std::endl;
    std::cout << System::String::Format(u"\tParent:\t{0} ({1})", args->get_NewParent()->get_NodeType(), System::ObjectExt::GetHashCode(args->get_NewParent())) << std::endl;
}

void ExNode::NodeChangingPrinter::NodeRemoving(System::SharedPtr<Aspose::Words::NodeChangingArgs> args)
{
    ASSERT_EQ(Aspose::Words::NodeChangingAction::Remove, args->get_Action());
}

void ExNode::NodeChangingPrinter::NodeRemoved(System::SharedPtr<Aspose::Words::NodeChangingArgs> args)
{
    ASSERT_EQ(Aspose::Words::NodeChangingAction::Remove, args->get_Action());
    ASSERT_TRUE(System::TestTools::IsNull(args->get_NewParent()));
    
    std::cout << System::String::Format(u"Removed node: {0} ({1})", args->get_Node()->get_NodeType(), System::ObjectExt::GetHashCode(args->get_Node())) << std::endl;
}


RTTI_INFO_IMPL_HASH(304439442u, ::Aspose::Words::ApiExamples::ExNode, ThisTypeBaseTypesInfo);

void ExNode::TraverseAllNodes(System::SharedPtr<Aspose::Words::CompositeNode> parentNode, int32_t depth)
{
    for (System::SharedPtr<Aspose::Words::Node> childNode = parentNode->get_FirstChild(); childNode != nullptr; childNode = childNode->get_NextSibling())
    {
        std::cout << System::String::Format(u"{0}{1}", System::String(u'\t', depth), Aspose::Words::Node::NodeTypeToString(childNode->get_NodeType()));
        
        // Recurse into the node if it is a composite node. Otherwise, print its contents if it is an inline node.
        if (childNode->get_IsComposite())
        {
            std::cout << std::endl;
            TraverseAllNodes(System::ExplicitCast<Aspose::Words::CompositeNode>(childNode), depth + 1);
        }
        else if (System::ObjectExt::Is<Aspose::Words::Inline>(childNode))
        {
            std::cout << System::String::Format(u" - \"{0}\"", childNode->GetText().Trim()) << std::endl;
        }
        else
        {
            std::cout << std::endl;
        }
    }
}

void ExNode::MapDocument(System::SharedPtr<System::Xml::XPath::XPathNavigator> navigator, System::SharedPtr<System::Text::StringBuilder> stringBuilder, int32_t depth)
{
    do
    {
        stringBuilder->Append(u' ', depth);
        stringBuilder->Append(navigator->get_Name() + u": ");
        
        if (navigator->get_Name() == u"Run")
        {
            stringBuilder->Append(navigator->get_Value());
        }
        
        stringBuilder->Append(u'\n');
        
        if (navigator->get_HasChildren())
        {
            navigator->MoveToFirstChild();
            MapDocument(navigator, stringBuilder, depth + 1);
            navigator->MoveToParent();
        }
    } while (navigator->MoveToNext());
}

void ExNode::TestNodeXPathNavigator(System::String navigatorResult, System::SharedPtr<Aspose::Words::Document> doc)
{
    for (auto&& run : System::IterateOver(doc->GetChildNodes(Aspose::Words::NodeType::Run, true)->ToArray()->LINQ_OfType<System::SharedPtr<Aspose::Words::Run> >()))
    {
        ASSERT_TRUE(navigatorResult.Contains(run->GetText().Trim()));
    }
}


namespace gtest_test
{

class ExNode : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExNode> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExNode>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExNode> ExNode::s_instance;

} // namespace gtest_test

void ExNode::CloneCompositeNode()
{
    //ExStart
    //ExFor:Node
    //ExFor:Node.Clone
    //ExSummary:Shows how to clone a composite node.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!"));
    
    // Below are two ways of cloning a composite node.
    // 1 -  Create a clone of a node, and create a clone of each of its child nodes as well.
    System::SharedPtr<Aspose::Words::Node> cloneWithChildren = System::ExplicitCast<Aspose::Words::Node>(para)->Clone(true);
    
    ASSERT_TRUE((System::ExplicitCast<Aspose::Words::CompositeNode>(cloneWithChildren))->get_HasChildNodes());
    ASSERT_EQ(u"Hello world!", cloneWithChildren->GetText().Trim());
    
    // 2 -  Create a clone of a node just by itself without any children.
    System::SharedPtr<Aspose::Words::Node> cloneWithoutChildren = System::ExplicitCast<Aspose::Words::Node>(para)->Clone(false);
    
    ASSERT_FALSE((System::ExplicitCast<Aspose::Words::CompositeNode>(cloneWithoutChildren))->get_HasChildNodes());
    ASSERT_EQ(System::String::Empty, cloneWithoutChildren->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExNode, CloneCompositeNode)
{
    s_instance->CloneCompositeNode();
}

} // namespace gtest_test

void ExNode::GetParentNode()
{
    //ExStart
    //ExFor:Node.ParentNode
    //ExSummary:Shows how to access a node's parent node.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
    
    // Append a child Run node to the document's first paragraph.
    auto run = System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!");
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    // The paragraph is the parent node of the run node. We can trace this lineage
    // all the way to the document node, which is the root of the document's node tree.
    ASPOSE_ASSERT_EQ(para, run->get_ParentNode());
    ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body(), para->get_ParentNode());
    ASPOSE_ASSERT_EQ(doc->get_FirstSection(), doc->get_FirstSection()->get_Body()->get_ParentNode());
    ASPOSE_ASSERT_EQ(doc, doc->get_FirstSection()->get_ParentNode());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExNode, GetParentNode)
{
    s_instance->GetParentNode();
}

} // namespace gtest_test

void ExNode::OwnerDocument()
{
    //ExStart
    //ExFor:Node.Document
    //ExFor:Node.ParentNode
    //ExSummary:Shows how to create a node and set its owning document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto para = System::MakeObject<Aspose::Words::Paragraph>(doc);
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!"));
    
    // We have not yet appended this paragraph as a child to any composite node.
    ASSERT_TRUE(System::TestTools::IsNull(para->get_ParentNode()));
    
    // If a node is an appropriate child node type of another composite node,
    // we can attach it as a child only if both nodes have the same owner document.
    // The owner document is the document we passed to the node's constructor.
    // We have not attached this paragraph to the document, so the document does not contain its text.
    ASPOSE_ASSERT_EQ(para->get_Document(), doc);
    ASSERT_EQ(System::String::Empty, doc->GetText().Trim());
    
    // Since the document owns this paragraph, we can apply one of its styles to the paragraph's contents.
    para->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
    
    // Add this node to the document, and then verify its contents.
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(para);
    
    ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body(), para->get_ParentNode());
    ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
    //ExEnd
    
    ASPOSE_ASSERT_EQ(doc, para->get_Document());
    ASSERT_FALSE(System::TestTools::IsNull(para->get_ParentNode()));
}

namespace gtest_test
{

TEST_F(ExNode, OwnerDocument)
{
    s_instance->OwnerDocument();
}

} // namespace gtest_test

void ExNode::ChildNodesEnumerate()
{
    //ExStart
    //ExFor:Node
    //ExFor:Node.CustomNodeId
    //ExFor:NodeType
    //ExFor:CompositeNode
    //ExFor:CompositeNode.GetChild
    //ExFor:CompositeNode.GetChildNodes(NodeType, bool)
    //ExFor:NodeCollection.Count
    //ExFor:NodeCollection.Item
    //ExSummary:Shows how to traverse through a composite node's collection of child nodes.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Add two runs and one shape as child nodes to the first paragraph of this document.
    auto paragraph = System::ExplicitCast<Aspose::Words::Paragraph>(doc->GetChild(Aspose::Words::NodeType::Paragraph, 0, true));
    paragraph->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"Hello world! "));
    
    auto shape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Rectangle);
    shape->set_Width(200);
    shape->set_Height(200);
    // Note that the 'CustomNodeId' is not saved to an output file and exists only during the node lifetime.
    shape->set_CustomNodeId(100);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::Inline);
    paragraph->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(shape);
    
    paragraph->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"Hello again!"));
    
    // Iterate through the paragraph's collection of immediate children,
    // and print any runs or shapes that we find within.
    System::SharedPtr<Aspose::Words::NodeCollection> children = paragraph->GetChildNodes(Aspose::Words::NodeType::Any, false);
    
    ASSERT_EQ(3, paragraph->GetChildNodes(Aspose::Words::NodeType::Any, false)->get_Count());
    
    for (auto&& child : System::IterateOver(children))
    {
        switch (child->get_NodeType())
        {
            case Aspose::Words::NodeType::Run:
                std::cout << "Run contents:" << std::endl;
                std::cout << System::String::Format(u"\t\"{0}\"", child->GetText().Trim()) << std::endl;
                break;
            
            case Aspose::Words::NodeType::Shape:
            {
                auto childShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(child);
                std::cout << "Shape:" << std::endl;
                std::cout << System::String::Format(u"\t{0}, {1}x{2}", childShape->get_ShapeType(), childShape->get_Width(), childShape->get_Height()) << std::endl;
                ASSERT_EQ(100, shape->get_CustomNodeId());
                break;
            }
            
            default:
                break;
        }
    }
    //ExEnd
    
    ASSERT_EQ(Aspose::Words::NodeType::Run, paragraph->GetChild(Aspose::Words::NodeType::Run, 0, true)->get_NodeType());
    ASSERT_EQ(u"Hello world! Hello again!", doc->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExNode, ChildNodesEnumerate)
{
    s_instance->ChildNodesEnumerate();
}

} // namespace gtest_test

void ExNode::RecurseChildren()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Paragraphs.docx");
    
    // Any node that can contain child nodes, such as the document itself, is composite.
    ASSERT_TRUE(doc->get_IsComposite());
    
    // Invoke the recursive function that will go through and print all the child nodes of a composite node.
    TraverseAllNodes(doc, 0);
}

namespace gtest_test
{

TEST_F(ExNode, RecurseChildren)
{
    s_instance->RecurseChildren();
}

} // namespace gtest_test

void ExNode::RemoveNodes()
{
    
    //ExStart
    //ExFor:Node
    //ExFor:Node.NodeType
    //ExFor:Node.Remove
    //ExSummary:Shows how to remove all child nodes of a specific type from a composite node.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
    
    System::SharedPtr<Aspose::Words::Node> curNode = doc->get_FirstSection()->get_Body()->get_FirstChild();
    
    while (curNode != nullptr)
    {
        // Save the next sibling node as a variable in case we want to move to it after deleting this node.
        System::SharedPtr<Aspose::Words::Node> nextNode = curNode->get_NextSibling();
        
        // A section body can contain Paragraph and Table nodes.
        // If the node is a Table, remove it from the parent.
        if (curNode->get_NodeType() == Aspose::Words::NodeType::Table)
        {
            curNode->Remove();
        }
        
        curNode = nextNode;
    }
    
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExNode, RemoveNodes)
{
    s_instance->RemoveNodes();
}

} // namespace gtest_test

void ExNode::EnumNextSibling()
{
    //ExStart
    //ExFor:CompositeNode.FirstChild
    //ExFor:Node.NextSibling
    //ExFor:Node.NodeTypeToString
    //ExFor:Node.NodeType
    //ExSummary:Shows how to use a node's NextSibling property to enumerate through its immediate children.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Paragraphs.docx");
    
    for (System::SharedPtr<Aspose::Words::Node> node = doc->get_FirstSection()->get_Body()->get_FirstChild(); node != nullptr; node = node->get_NextSibling())
    {
        std::cout << std::endl;
        std::cout << System::String::Format(u"Node type: {0}", Aspose::Words::Node::NodeTypeToString(node->get_NodeType())) << std::endl;
        
        System::String contents = node->GetText().Trim();
        std::cout << (contents == System::String::Empty ? u"This node contains no text" : System::String::Format(u"Contents: \"{0}\"", node->GetText().Trim())) << std::endl;
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExNode, EnumNextSibling)
{
    s_instance->EnumNextSibling();
}

} // namespace gtest_test

void ExNode::TypedAccess()
{
    
    //ExStart
    //ExFor:Story.Tables
    //ExFor:Table.FirstRow
    //ExFor:Table.LastRow
    //ExFor:TableCollection
    //ExSummary:Shows how to remove the first and last rows of all tables in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    
    System::SharedPtr<Aspose::Words::Tables::TableCollection> tables = doc->get_FirstSection()->get_Body()->get_Tables();
    
    ASSERT_EQ(5, tables->idx_get(0)->get_Rows()->get_Count());
    ASSERT_EQ(4, tables->idx_get(1)->get_Rows()->get_Count());
    
    for (auto&& table : System::IterateOver(tables->LINQ_OfType<System::SharedPtr<Aspose::Words::Tables::Table> >()))
    {
        System::SharedPtr<Aspose::Words::Tables::Row> condExpression = table->get_FirstRow();
        if (condExpression != nullptr)
        {
            condExpression->Remove();
        }
        System::SharedPtr<Aspose::Words::Tables::Row> condExpression2 = table->get_LastRow();
        if (condExpression2 != nullptr)
        {
            condExpression2->Remove();
        }
    }
    
    ASSERT_EQ(3, tables->idx_get(0)->get_Rows()->get_Count());
    ASSERT_EQ(2, tables->idx_get(1)->get_Rows()->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExNode, TypedAccess)
{
    s_instance->TypedAccess();
}

} // namespace gtest_test

void ExNode::RemoveChild()
{
    //ExStart
    //ExFor:CompositeNode.LastChild
    //ExFor:Node.PreviousSibling
    //ExFor:CompositeNode.RemoveChild``1(``0)
    //ExSummary:Shows how to use of methods of Node and CompositeNode to remove a section before the last section in the document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Section 1 text.");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakContinuous);
    builder->Writeln(u"Section 2 text.");
    
    // Both sections are siblings of each other.
    auto lastSection = System::ExplicitCast<Aspose::Words::Section>(doc->get_LastChild());
    auto firstSection = System::ExplicitCast<Aspose::Words::Section>(lastSection->get_PreviousSibling());
    
    // Remove a section based on its sibling relationship with another section.
    if (lastSection->get_PreviousSibling() != nullptr)
    {
        doc->RemoveChild<System::SharedPtr<Aspose::Words::Section>>(firstSection);
    }
    
    // The section we removed was the first one, leaving the document with only the second.
    ASSERT_EQ(u"Section 2 text.", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExNode, RemoveChild)
{
    s_instance->RemoveChild();
}

} // namespace gtest_test

void ExNode::CreateAndAddParagraphNode()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto para = System::MakeObject<Aspose::Words::Paragraph>(doc);
    
    System::SharedPtr<Aspose::Words::Section> section = doc->get_LastSection();
    section->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(para);
}

namespace gtest_test
{

TEST_F(ExNode, CreateAndAddParagraphNode)
{
    s_instance->CreateAndAddParagraphNode();
}

} // namespace gtest_test

void ExNode::RemoveSmartTagsFromCompositeNode()
{
    //ExStart
    //ExFor:CompositeNode.RemoveSmartTags
    //ExSummary:Removes all smart tags from descendant nodes of a composite node.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Smart tags.doc");
    
    ASSERT_EQ(8, doc->GetChildNodes(Aspose::Words::NodeType::SmartTag, true)->get_Count());
    
    doc->RemoveSmartTags();
    
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::SmartTag, true)->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExNode, RemoveSmartTagsFromCompositeNode)
{
    s_instance->RemoveSmartTagsFromCompositeNode();
}

} // namespace gtest_test

void ExNode::GetIndexOfNode()
{
    //ExStart
    //ExFor:CompositeNode.IndexOf
    //ExSummary:Shows how to get the index of a given child node from its parent.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    System::SharedPtr<Aspose::Words::Body> body = doc->get_FirstSection()->get_Body();
    
    // Retrieve the index of the last paragraph in the body of the first section.
    ASSERT_EQ(24, body->GetChildNodes(Aspose::Words::NodeType::Any, false)->IndexOf(body->get_LastParagraph()));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExNode, GetIndexOfNode)
{
    s_instance->GetIndexOfNode();
}

} // namespace gtest_test

void ExNode::ConvertNodeToHtmlWithDefaultOptions()
{
    //ExStart
    //ExFor:Node.ToString(SaveFormat)
    //ExFor:Node.ToString(SaveOptions)
    //ExSummary:Exports the content of a node to String in HTML format.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    System::SharedPtr<Aspose::Words::Node> node = doc->get_LastSection()->get_Body()->get_LastParagraph();
    
    // When we call the ToString method using the html SaveFormat overload,
    // it converts the node's contents to their raw html representation.
    ASSERT_EQ(System::String(u"<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%; font-size:12pt\">") + u"<span style=\"font-family:'Times New Roman'\">Hello World!</span>" + u"</p>", node->ToString(Aspose::Words::SaveFormat::Html));
    
    // We can also modify the result of this conversion using a SaveOptions object.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_ExportRelativeFontSize(true);
    
    ASSERT_EQ(System::String(u"<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%\">") + u"<span style=\"font-family:'Times New Roman'\">Hello World!</span>" + u"</p>", node->ToString(saveOptions));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExNode, ConvertNodeToHtmlWithDefaultOptions)
{
    s_instance->ConvertNodeToHtmlWithDefaultOptions();
}

} // namespace gtest_test

void ExNode::TypedNodeCollectionToArray()
{
    //ExStart
    //ExFor:ParagraphCollection.ToArray
    //ExSummary:Shows how to create an array from a NodeCollection.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Paragraphs.docx");
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Paragraph>> paras = doc->get_FirstSection()->get_Body()->get_Paragraphs()->ToArray();
    
    ASSERT_EQ(22, paras->get_Length());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExNode, TypedNodeCollectionToArray)
{
    s_instance->TypedNodeCollectionToArray();
}

} // namespace gtest_test

void ExNode::NodeEnumerationHotRemove()
{
    //ExStart
    //ExFor:ParagraphCollection.ToArray
    //ExSummary:Shows how to use "hot remove" to remove a node during enumeration.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"The first paragraph");
    builder->Writeln(u"The second paragraph");
    builder->Writeln(u"The third paragraph");
    builder->Writeln(u"The fourth paragraph");
    
    // Remove a node from the collection in the middle of an enumeration.
    for (System::SharedPtr<Aspose::Words::Paragraph> para : doc->get_FirstSection()->get_Body()->get_Paragraphs()->ToArray())
    {
        if (para->get_Range()->get_Text().Contains(u"third"))
        {
            para->Remove();
        }
    }
    
    
    ASSERT_FALSE(doc->GetText().Contains(u"The third paragraph"));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExNode, NodeEnumerationHotRemove)
{
    s_instance->NodeEnumerationHotRemove();
}

} // namespace gtest_test

void ExNode::NodeChangingCallback()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->set_NodeChangingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExNode::NodeChangingPrinter>());
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Cell 1");
    builder->InsertCell();
    builder->Write(u"Cell 2");
    builder->EndTable();
    
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    builder->get_CurrentParagraph()->get_ParentNode()->RemoveAllChildren();
}

namespace gtest_test
{

TEST_F(ExNode, NodeChangingCallback)
{
    s_instance->NodeChangingCallback();
}

} // namespace gtest_test

void ExNode::NodeCollection()
{
    //ExStart
    //ExFor:NodeCollection.Contains(Node)
    //ExFor:NodeCollection.Insert(Int32,Node)
    //ExFor:NodeCollection.Remove(Node)
    //ExSummary:Shows how to work with a NodeCollection.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Add text to the document by inserting Runs using a DocumentBuilder.
    builder->Write(u"Run 1. ");
    builder->Write(u"Run 2. ");
    
    // Every invocation of the "Write" method creates a new Run,
    // which then appears in the parent Paragraph's RunCollection.
    System::SharedPtr<Aspose::Words::RunCollection> runs = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs();
    
    ASSERT_EQ(2, runs->get_Count());
    
    // We can also insert a node into the RunCollection manually.
    auto newRun = System::MakeObject<Aspose::Words::Run>(doc, u"Run 3. ");
    runs->Insert(3, newRun);
    
    ASSERT_TRUE(runs->Contains(newRun));
    ASSERT_EQ(u"Run 1. Run 2. Run 3.", doc->GetText().Trim());
    
    // Access individual runs and remove them to remove their text from the document.
    System::SharedPtr<Aspose::Words::Run> run = runs->idx_get(1);
    runs->Remove(run);
    
    ASSERT_EQ(u"Run 1. Run 3.", doc->GetText().Trim());
    ASSERT_FALSE(System::TestTools::IsNull(run));
    ASSERT_FALSE(runs->Contains(run));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExNode, NodeCollection)
{
    s_instance->NodeCollection();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
