#include "stdafx.h"

#include "Common.h"

#include <system/enumerator_adapter.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/Text/Paragraph.h>
#include <Model/Styles/Style.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/Body.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/CompositeNode.h>
#include <Model/Importing/NodeImporter.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/DocumentBase.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

namespace
{
    // ExStart:CommonExtractContentHelperMethods
    System::SharedPtr<Node> GetAncestorInBody(System::SharedPtr<Node> startNode)
    {
        while (startNode->get_ParentNode()->get_NodeType() != NodeType::Body)
        {
            startNode = startNode->get_ParentNode();
        }
        return startNode;
    }

    std::vector<System::SharedPtr<Node>> FillSelfAndParents(System::SharedPtr<Node> node, System::SharedPtr<Node> tillNode)
    {
        std::vector<System::SharedPtr<Node>> dest;
        System::SharedPtr<Node> currentNode = node;
        while (currentNode != tillNode)
        {
            dest.push_back(currentNode);
            currentNode = currentNode->get_ParentNode();
        }

        return dest;
    }

    void RemoveNodesOutsideOfRange(System::SharedPtr<Node> markerNode, bool isInclusive, bool isStartMarker)
    {
        // Remove the nodes up to/from the marker.
        bool isSkip = false;
        bool isProcessing = true;
        bool isRemoving = isStartMarker;
        System::SharedPtr<Node> nextNode = markerNode->get_ParentNode()->get_FirstChild();

        while (isProcessing && nextNode != nullptr)
        {
            System::SharedPtr<Node> currentNode = nextNode;
            isSkip = false;

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

    void VerifyParameterNodes(const System::SharedPtr<Node>& startNode, const System::SharedPtr<Node>& endNode)
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
        // Check the end node is after the start node in the DOM tree
        // First check if they are in different sections, then if they're not check their position in the body of the same section they are in.
        System::SharedPtr<Section> startSection = System::DynamicCast<Section>(startNode->GetAncestor(NodeType::Section));
        System::SharedPtr<Section> endSection = System::DynamicCast<Section>(endNode->GetAncestor(NodeType::Section));
        int32_t startIndex = startSection->get_ParentNode()->IndexOf(startSection);
        int32_t endIndex = endSection->get_ParentNode()->IndexOf(endSection);
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

    System::SharedPtr<Node> FindNextNode(NodeType nodeType, System::SharedPtr<Node> fromNode)
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

    void ProcessMarker(System::SharedPtr<Node> cloneNode,
                       std::vector<System::SharedPtr<Node>> &nodes,
                       System::SharedPtr<Node> node,
                       System::SharedPtr<Node> blockLevelAncestor,
                       bool isInclusive,
                       bool isStartMarker,
                       bool canAdd,
                       bool forceAdd)
    {
        // If we are dealing with a block level node just see if it should be included and add it to the list.
        if (node == blockLevelAncestor)
        {
            if (canAdd && isInclusive)
            {
                nodes.push_back(cloneNode);
            }
            return;
        }

        // cloneNode is a clone of blockLevelNode. If node != blockLevelNode, blockLevelAncestor is ancestor of node; that means it is a composite node.
        System::Diagnostics::Debug::Assert(cloneNode->get_IsComposite());

        // If a marker is a FieldStart node check if it's to be included or not.
        // We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
        if (node->get_NodeType() == NodeType::FieldStart)
        {
            // If the marker is a start node and is not be included then skip to the end of the field.
            // If the marker is an end node and it is to be included then move to the end field so the field will not be removed.
            if ((isStartMarker && !isInclusive) || (!isStartMarker && isInclusive))
            {
                while (node->get_NextSibling() != nullptr && node->get_NodeType() != NodeType::FieldEnd)
                {
                    node = node->get_NextSibling();
                }
            }
        }

        // Support a case if marker node is on the third level of document body or lower.
        std::vector<System::SharedPtr<Node>> nodeBranch(FillSelfAndParents(node, blockLevelAncestor));

        // Process the corresponding node in our cloned node by index.
        System::SharedPtr<Node> currentCloneNode = cloneNode;
        for (int index = nodeBranch.size() - 1; index >= 0; --index)
        {
            System::SharedPtr<Node> currentNode = nodeBranch.at(index);
            int32_t nodeIndex = currentNode->get_ParentNode()->IndexOf(currentNode);
            currentCloneNode = (System::DynamicCast<CompositeNode>(currentCloneNode))->get_ChildNodes()->idx_get(nodeIndex);
            RemoveNodesOutsideOfRange(currentCloneNode, isInclusive || (index > 0), isStartMarker);
        }

        // After processing the composite node may become empty. If it has don't include it.
        if (canAdd && (forceAdd || (System::DynamicCast<CompositeNode>(cloneNode))->get_HasChildNodes()))
        {
            nodes.push_back(cloneNode);
        }
    }

    void IncludeNextParagraph(System::SharedPtr<Node> node, std::vector<System::SharedPtr<Node>> &nodes)
    {
        System::SharedPtr<Paragraph> paragraph = System::DynamicCast<Paragraph>(FindNextNode(NodeType::Paragraph, node->get_NextSibling()));
        if (paragraph != nullptr)
        {
            // Move to first child to include paragraph without contents.
            System::SharedPtr<Node> markerNode = paragraph->get_HasChildNodes() ? paragraph->get_FirstChild() : paragraph;
            System::SharedPtr<Node> rootNode = GetAncestorInBody(paragraph);

            ProcessMarker(rootNode->Clone(true), nodes, markerNode, rootNode, markerNode == paragraph, false, true, true);
        }
    }
    // ExEnd:CommonExtractContentHelperMethods
}

// ExStart:CommonExtractContent
std::vector<System::SharedPtr<Node>> ExtractContent(System::SharedPtr<Node> startNode, System::SharedPtr<Node> endNode, bool isInclusive)
{
    // First check that the nodes passed to this method are valid for use.
    VerifyParameterNodes(startNode, endNode);

    // Create a list to store the extracted nodes.
    std::vector<System::SharedPtr<Node>> nodes;

    // If either marker is part of a comment then to include the comment itself we need to move the pointer
    // forward to the Comment Node found after the CommentRangeEnd node.
    if (endNode->get_NodeType() == NodeType::CommentRangeEnd && isInclusive)
    {
        System::SharedPtr<Node> node = FindNextNode(NodeType::Comment, endNode->get_NextSibling());
        if (node != nullptr)
        {
            endNode = node;
        }
    }

    // Keep a record of the original nodes passed to this method so we can split marker nodes if needed.
    System::SharedPtr<Node> originalStartNode = startNode;
    System::SharedPtr<Node> originalEndNode = endNode;

    // Extract content based on block level nodes (paragraphs and tables). Traverse through parent nodes to find them.
    // We will split the content of first and last nodes depending if the marker nodes are inline
    startNode = GetAncestorInBody(startNode);
    endNode = GetAncestorInBody(endNode);

    bool isExtracting = true;
    bool isStartingNode = true;
    bool isEndingNode = false;
    // The current node we are extracting from the document.
    System::SharedPtr<Node> currNode = startNode;

    // Begin extracting content. Process all block level nodes and specifically split the first and last nodes when needed so paragraph formatting is retained.
    // Method is little more complex than a regular extractor as we need to factor in extracting using inline nodes, fields, bookmarks etc as to make it really useful.
    while (isExtracting)
    {
        // Clone the current node and its children to obtain a copy.
        System::SharedPtr<Node> cloneNode = currNode->Clone(true);
        isEndingNode = System::ObjectExt::Equals(currNode, endNode);

        if (isStartingNode || isEndingNode)
        {
            // We need to process each marker separately so pass it off to a separate method instead.
            // End should be processed at first to keep node indexes.
            if (isEndingNode)
            {
                // !isStartingNode: don't add the node twice if the markers are the same node.
                ProcessMarker(cloneNode, nodes, originalEndNode, currNode, isInclusive, false, !isStartingNode, false);
                isExtracting = false;
            }

            // Conditional needs to be separate as the block level start and end markers maybe the same node.
            if (isStartingNode)
            {
                ProcessMarker(cloneNode, nodes, originalStartNode, currNode, isInclusive, true, true, false);
                isStartingNode = false;
            }
        }
        else
        {
            nodes.push_back(cloneNode);
        }

        // Move to the next node and extract it. If next node is null that means the rest of the content is found in a different section.
        if (currNode->get_NextSibling() == nullptr && isExtracting)
        {
            // Move to the next section.
            System::SharedPtr<Section> nextSection = System::DynamicCast<Section>(currNode->GetAncestor(NodeType::Section)->get_NextSibling());
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
// ExEnd:CommonExtractContent

std::vector<System::SharedPtr<Paragraph>> ParagraphsByStyleName(const System::SharedPtr<Document>& doc, const System::String& styleName)
{
    // Create an array to collect paragraphs of the specified style.
    std::vector<System::SharedPtr<Paragraph>> paragraphsWithStyle;
    // Get all paragraphs from the document.
    System::SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(NodeType::Paragraph, true);
    // Look through all paragraphs to find those with the specified style.
    for (System::SharedPtr<Paragraph> paragraph : System::IterateOver<System::SharedPtr<Paragraph>>(paragraphs))
    {
        if (paragraph->get_ParagraphFormat()->get_Style()->get_Name() == styleName)
        {
            paragraphsWithStyle.push_back(paragraph);
        }
    }

    return paragraphsWithStyle;
}

// ExStart:CommonGenerateDocument
System::SharedPtr<Document> GenerateDocument(const System::SharedPtr<Document>& srcDoc, const std::vector<System::SharedPtr<Node>>& nodes)
{
    // Create a blank document.
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>();
    // Remove the first paragraph from the empty document.
    dstDoc->get_FirstSection()->get_Body()->RemoveAllChildren();

    // Import each node from the list into the new document. Keep the original formatting of the node.
    System::SharedPtr<NodeImporter> importer = System::MakeObject<NodeImporter>(srcDoc, dstDoc, ImportFormatMode::KeepSourceFormatting);

    for (System::SharedPtr<Node> node : nodes)
    {
        System::SharedPtr<Node> importNode = importer->ImportNode(node, true);
        dstDoc->get_FirstSection()->get_Body()->AppendChild(importNode);
    }

    // Return the generated document.
    return dstDoc;
}
// ExEnd:CommonGenerateDocument