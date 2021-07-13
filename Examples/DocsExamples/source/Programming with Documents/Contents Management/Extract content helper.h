#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeImporter.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Style.h>
#include <system/collections/list.h>
#include <system/diagnostics/debug.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/object_ext.h>

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Management {

class ExtractContentHelper : public System::Object
{
public:
    static System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Node>>> ExtractContent(System::SharedPtr<Node> startNode,
                                                                                                         System::SharedPtr<Node> endNode, bool isInclusive)
    {
        // First, check that the nodes passed to this method are valid for use.
        VerifyParameterNodes(startNode, endNode);

        // Create a list to store the extracted nodes.
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Node>>> nodes =
            System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Node>>>();

        // If either marker is part of a comment, including the comment itself, we need to move the pointer
        // forward to the Comment Node found after the CommentRangeEnd node.
        if (endNode->get_NodeType() == NodeType::CommentRangeEnd && isInclusive)
        {
            System::SharedPtr<Node> node = FindNextNode(NodeType::Comment, endNode->get_NextSibling());
            if (node != nullptr)
            {
                endNode = node;
            }
        }

        // Keep a record of the original nodes passed to this method to split marker nodes if needed.
        System::SharedPtr<Node> originalStartNode = startNode;
        System::SharedPtr<Node> originalEndNode = endNode;

        // Extract content based on block-level nodes (paragraphs and tables). Traverse through parent nodes to find them.
        // We will split the first and last nodes' content, depending if the marker nodes are inline.
        startNode = GetAncestorInBody(startNode);
        endNode = GetAncestorInBody(endNode);

        bool isExtracting = true;
        bool isStartingNode = true;
        // The current node we are extracting from the document.
        System::SharedPtr<Node> currNode = startNode;

        // Begin extracting content. Process all block-level nodes and specifically split the first
        // and last nodes when needed, so paragraph formatting is retained.
        // Method is a little more complicated than a regular extractor as we need to factor
        // in extracting using inline nodes, fields, bookmarks, etc. to make it useful.
        while (isExtracting)
        {
            // Clone the current node and its children to obtain a copy.
            System::SharedPtr<Node> cloneNode = currNode->Clone(true);
            bool isEndingNode = System::ObjectExt::Equals(currNode, endNode);

            if (isStartingNode || isEndingNode)
            {
                // We need to process each marker separately, so pass it off to a separate method instead.
                // End should be processed at first to keep node indexes.
                if (isEndingNode)
                {
                    // !isStartingNode: don't add the node twice if the markers are the same node.
                    ProcessMarker(cloneNode, nodes, originalEndNode, currNode, isInclusive, false, !isStartingNode, false);
                    isExtracting = false;
                }

                // Conditional needs to be separate as the block level start and end markers, maybe the same node.
                if (isStartingNode)
                {
                    ProcessMarker(cloneNode, nodes, originalStartNode, currNode, isInclusive, true, true, false);
                    isStartingNode = false;
                }
            }
            else
            {
                nodes->Add(cloneNode);
            }

            // Move to the next node and extract it. If the next node is null,
            // the rest of the content is found in a different section.
            if (currNode->get_NextSibling() == nullptr && isExtracting)
            {
                // Move to the next section.
                auto nextSection = System::DynamicCast<Section>(currNode->GetAncestor(NodeType::Section)->get_NextSibling());
                currNode = nextSection->get_Body()->get_FirstChild();
            }
            else
            {
                // Move to the next node in the body.
                currNode = currNode->get_NextSibling();
            }
        }

        // For compatibility with mode with inline bookmarks, add the next paragraph (empty).
        if (isInclusive && originalEndNode == endNode && !originalEndNode->get_IsComposite())
        {
            IncludeNextParagraph(endNode, nodes);
        }

        // Return the nodes between the node markers.
        return nodes;
    }

    static System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Paragraph>>> ParagraphsByStyleName(System::SharedPtr<Document> doc,
                                                                                                                     System::String styleName)
    {
        // Create an array to collect paragraphs of the specified style.
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Paragraph>>> paragraphsWithStyle =
            System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Paragraph>>>();

        System::SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(NodeType::Paragraph, true);

        // Look through all paragraphs to find those with the specified style.
        for (const auto& paragraph : System::IterateOver<Paragraph>(paragraphs))
        {
            if (paragraph->get_ParagraphFormat()->get_Style()->get_Name() == styleName)
            {
                paragraphsWithStyle->Add(paragraph);
            }
        }

        return paragraphsWithStyle;
    }

    static System::SharedPtr<Document> GenerateDocument(System::SharedPtr<Document> srcDoc,
                                                        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Node>>> nodes)
    {
        auto dstDoc = System::MakeObject<Document>();
        // Remove the first paragraph from the empty document.
        dstDoc->get_FirstSection()->get_Body()->RemoveAllChildren();

        // Import each node from the list into the new document. Keep the original formatting of the node.
        auto importer = System::MakeObject<NodeImporter>(srcDoc, dstDoc, ImportFormatMode::KeepSourceFormatting);

        for (const auto& node : nodes)
        {
            System::SharedPtr<Node> importNode = importer->ImportNode(node, true);
            dstDoc->get_FirstSection()->get_Body()->AppendChild(importNode);
        }

        return dstDoc;
    }

private:
    static void VerifyParameterNodes(System::SharedPtr<Node> startNode, System::SharedPtr<Node> endNode)
    {
        // The order in which these checks are done is important.
        if (startNode == nullptr)
        {
            throw System::ArgumentException(u"Start node cannot be null");
        }
        if (endNode == nullptr)
        {
            throw System::ArgumentException(u"End node cannot be null");
        }

        if (!System::ObjectExt::Equals(startNode->get_Document(), endNode->get_Document()))
        {
            throw System::ArgumentException(u"Start node and end node must belong to the same document");
        }

        if (startNode->GetAncestor(NodeType::Body) == nullptr || endNode->GetAncestor(NodeType::Body) == nullptr)
        {
            throw System::ArgumentException(u"Start node and end node must be a child or descendant of a body");
        }

        // Check the end node is after the start node in the DOM tree.
        // First, check if they are in different sections, then if they're not,
        // check their position in the body of the same section.
        auto startSection = System::DynamicCast<Section>(startNode->GetAncestor(NodeType::Section));
        auto endSection = System::DynamicCast<Section>(endNode->GetAncestor(NodeType::Section));

        int startIndex = startSection->get_ParentNode()->IndexOf(startSection);
        int endIndex = endSection->get_ParentNode()->IndexOf(endSection);

        if (startIndex == endIndex)
        {
            if (startSection->get_Body()->IndexOf(GetAncestorInBody(startNode)) > endSection->get_Body()->IndexOf(GetAncestorInBody(endNode)))
            {
                throw System::ArgumentException(u"The end node must be after the start node in the body");
            }
        }
        else if (startIndex > endIndex)
        {
            throw System::ArgumentException(u"The section of end node must be after the section start node");
        }
    }

    static System::SharedPtr<Node> FindNextNode(NodeType nodeType, System::SharedPtr<Node> fromNode)
    {
        if (fromNode == nullptr || fromNode->get_NodeType() == nodeType)
        {
            return fromNode;
        }

        if (fromNode->get_IsComposite())
        {
            System::SharedPtr<Node> node = FindNextNode(nodeType, (System::DynamicCast<CompositeNode>(fromNode))->get_FirstChild());
            if (node != nullptr)
            {
                return node;
            }
        }

        return FindNextNode(nodeType, fromNode->get_NextSibling());
    }

    bool IsInline(System::SharedPtr<Node> node)
    {
        // Test if the node is a descendant of a Paragraph or Table node and is not a paragraph
        // or a table a paragraph inside a comment class that is decent of a paragraph is possible.
        return (node->GetAncestor(NodeType::Paragraph) != nullptr || node->GetAncestor(NodeType::Table) != nullptr) &&
               !(node->get_NodeType() == NodeType::Paragraph || node->get_NodeType() == NodeType::Table);
    }

    static void ProcessMarker(System::SharedPtr<Node> cloneNode, System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Node>>> nodes,
                              System::SharedPtr<Node> node, System::SharedPtr<Node> blockLevelAncestor, bool isInclusive, bool isStartMarker, bool canAdd,
                              bool forceAdd)
    {
        // If we are dealing with a block-level node, see if it should be included and add it to the list.
        if (node == blockLevelAncestor)
        {
            if (canAdd && isInclusive)
            {
                nodes->Add(cloneNode);
            }
            return;
        }

        // cloneNode is a clone of blockLevelNode. If node != blockLevelNode, blockLevelAncestor
        // is the node's ancestor that means it is a composite node.
        System::Diagnostics::Debug::Assert(cloneNode->get_IsComposite());

        // If a marker is a FieldStart node check if it's to be included or not.
        // We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
        if (node->get_NodeType() == NodeType::FieldStart)
        {
            // If the marker is a start node and is not included, skip to the end of the field.
            // If the marker is an end node and is to be included, then move to the end field so the field will not be removed.
            if ((isStartMarker && !isInclusive) || (!isStartMarker && isInclusive))
            {
                while (node->get_NextSibling() != nullptr && node->get_NodeType() != NodeType::FieldEnd)
                {
                    node = node->get_NextSibling();
                }
            }
        }

        // Support a case if the marker node is on the third level of the document body or lower.
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Node>>> nodeBranch = FillSelfAndParents(node, blockLevelAncestor);

        // Process the corresponding node in our cloned node by index.
        System::SharedPtr<Node> currentCloneNode = cloneNode;
        for (int i = nodeBranch->get_Count() - 1; i >= 0; i--)
        {
            System::SharedPtr<Node> currentNode = nodeBranch->idx_get(i);
            int nodeIndex = currentNode->get_ParentNode()->IndexOf(currentNode);
            currentCloneNode = (System::DynamicCast<CompositeNode>(currentCloneNode))->get_ChildNodes()->idx_get(nodeIndex);

            RemoveNodesOutsideOfRange(currentCloneNode, isInclusive || (i > 0), isStartMarker);
        }

        // After processing, the composite node may become empty if it has doesn't include it.
        if (canAdd && (forceAdd || (System::DynamicCast<CompositeNode>(cloneNode))->get_HasChildNodes()))
        {
            nodes->Add(cloneNode);
        }
    }

    static void RemoveNodesOutsideOfRange(System::SharedPtr<Node> markerNode, bool isInclusive, bool isStartMarker)
    {
        bool isProcessing = true;
        bool isRemoving = isStartMarker;
        System::SharedPtr<Node> nextNode = markerNode->get_ParentNode()->get_FirstChild();

        while (isProcessing && nextNode != nullptr)
        {
            System::SharedPtr<Node> currentNode = nextNode;
            bool isSkip = false;

            if (System::ObjectExt::Equals(currentNode, markerNode))
            {
                if (isStartMarker)
                {
                    isProcessing = false;
                    if (isInclusive)
                    {
                        isRemoving = false;
                    }
                }
                else
                {
                    isRemoving = true;
                    if (isInclusive)
                    {
                        isSkip = true;
                    }
                }
            }

            nextNode = nextNode->get_NextSibling();
            if (isRemoving && !isSkip)
            {
                currentNode->Remove();
            }
        }
    }

    static System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Node>>> FillSelfAndParents(System::SharedPtr<Node> node,
                                                                                                             System::SharedPtr<Node> tillNode)
    {
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Node>>> list =
            System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Node>>>();
        System::SharedPtr<Node> currentNode = node;

        while (currentNode != tillNode)
        {
            list->Add(currentNode);
            currentNode = currentNode->get_ParentNode();
        }

        return list;
    }

    static void IncludeNextParagraph(System::SharedPtr<Node> node, System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Node>>> nodes)
    {
        auto paragraph = System::DynamicCast<Paragraph>(FindNextNode(NodeType::Paragraph, node->get_NextSibling()));
        if (paragraph != nullptr)
        {
            // Move to the first child to include paragraphs without content.
            System::SharedPtr<Node> markerNode = paragraph->get_HasChildNodes() ? paragraph->get_FirstChild() : paragraph;
            System::SharedPtr<Node> rootNode = GetAncestorInBody(paragraph);

            ProcessMarker(rootNode->Clone(true), nodes, markerNode, rootNode, markerNode == paragraph, false, true, true);
        }
    }

    static System::SharedPtr<Node> GetAncestorInBody(System::SharedPtr<Node> startNode)
    {
        while (startNode->get_ParentNode()->get_NodeType() != NodeType::Body)
        {
            startNode = startNode->get_ParentNode();
        }
        return startNode;
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Contents_Management
