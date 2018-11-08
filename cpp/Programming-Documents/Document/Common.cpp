#include "stdafx.h"
#include "Common.h"

#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/collections/list.h>
#include <system/collections/ilist.h>
#include <system/collections/ienumerator.h>
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
#include <cstdint>

using namespace Aspose::Words;

namespace
{
    // ExStart:CommonExtractContentHelperMethods
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
        if (startNode->GetAncestor(Aspose::Words::NodeType::Body) == nullptr || endNode->GetAncestor(Aspose::Words::NodeType::Body) == nullptr)
        {
            throw System::ArgumentException(u"Start node and end node must be a child or descendant of a body");
        }
        // Check the end node is after the start node in the DOM tree
        // First check if they are in different sections, then if they're not check their position in the body of the same section they are in.
        System::SharedPtr<Section> startSection = System::DynamicCast<Section>(startNode->GetAncestor(Aspose::Words::NodeType::Section));
        System::SharedPtr<Section> endSection = System::DynamicCast<Section>(endNode->GetAncestor(Aspose::Words::NodeType::Section));
        int32_t startIndex = startSection->get_ParentNode()->IndexOf(startSection);
        int32_t endIndex = endSection->get_ParentNode()->IndexOf(endSection);
        if (startIndex == endIndex)
        {
            if (startSection->get_Body()->IndexOf(startNode) > endSection->get_Body()->IndexOf(endNode))
            {
                throw System::ArgumentException(u"The end node must be after the start node in the body");
            }
        }
        else if (startIndex > endIndex)
        {
            throw System::ArgumentException(u"The section of end node must be after the section start node");
        }
    }

    bool IsInline(const System::SharedPtr<Node>& node)
    {
        // Test if the node is desendant of a Paragraph or Table node and also is not a paragraph or a table a paragraph inside a comment class which is decesant of a pararaph is possible.
        return ((node->GetAncestor(Aspose::Words::NodeType::Paragraph) != nullptr || node->GetAncestor(Aspose::Words::NodeType::Table) != nullptr)
            && !(node->get_NodeType() == Aspose::Words::NodeType::Paragraph || node->get_NodeType() == Aspose::Words::NodeType::Table));
    }

    void ProcessMarker(const System::SharedPtr<CompositeNode>& cloneNode, std::vector<System::SharedPtr<Node>>& nodes, System::SharedPtr<Node> node, bool isInclusive, bool isStartMarker, bool isEndMarker)
    {
        // If we are dealing with a block level node just see if it should be included and add it to the list.
        if (!IsInline(node))
        {
            // Don't add the node twice if the markers are the same node
            if (!(isStartMarker && isEndMarker))
            {
                if (isInclusive)
                {
                    nodes.push_back(cloneNode);
                }
            }
            return;
        }

        // If a marker is a FieldStart node check if it's to be included or not.
        // We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
        if (node->get_NodeType() == Aspose::Words::NodeType::FieldStart)
        {
            // If the marker is a start node and is not be included then skip to the end of the field.
            // If the marker is an end node and it is to be included then move to the end field so the field will not be removed.
            if ((isStartMarker && !isInclusive) || (!isStartMarker && isInclusive))
            {
                while (node->get_NextSibling() != nullptr && node->get_NodeType() != Aspose::Words::NodeType::FieldEnd)
                {
                    node = node->get_NextSibling();
                }
            }
        }

        // If either marker is part of a comment then to include the comment itself we need to move the pointer forward to the Comment
        // Node found after the CommentRangeEnd node.
        if (node->get_NodeType() == Aspose::Words::NodeType::CommentRangeEnd)
        {
            while (node->get_NextSibling() != nullptr && node->get_NodeType() != Aspose::Words::NodeType::Comment)
            {
                node = node->get_NextSibling();
            }
        }

        // Find the corresponding node in our cloned node by index and return it.
        // If the start and end node are the same some child nodes might already have been removed. Subtract the
        // Difference to get the right index.
        int32_t indexDiff = node->get_ParentNode()->get_ChildNodes()->get_Count() - cloneNode->get_ChildNodes()->get_Count();

        // Child node count identical.
        if (indexDiff == 0)
        {
            node = cloneNode->get_ChildNodes()->idx_get(node->get_ParentNode()->IndexOf(node));
        }
        else
        {
            node = cloneNode->get_ChildNodes()->idx_get(node->get_ParentNode()->IndexOf(node) - indexDiff);
        }

        // Remove the nodes up to/from the marker.
        bool isSkip = false;
        bool isProcessing = true;
        bool isRemoving = isStartMarker;
        System::SharedPtr<Node> nextNode = cloneNode->get_FirstChild();

        while (isProcessing && nextNode != nullptr)
        {
            System::SharedPtr<Node> currentNode = nextNode;
            isSkip = false;
            if (System::ObjectExt::Equals(currentNode, node))
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

        // After processing the composite node may become empty. If it has don't include it.
        if (!(isStartMarker && isEndMarker))
        {
            if (cloneNode->get_HasChildNodes())
            {
                nodes.push_back(cloneNode);
            }
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

    // Keep a record of the original nodes passed to this method so we can split marker nodes if needed.
    System::SharedPtr<Node> originalStartNode = startNode;
    System::SharedPtr<Node> originalEndNode = endNode;

    // Extract content based on block level nodes (paragraphs and tables). Traverse through parent nodes to find them.
    // We will split the content of first and last nodes depending if the marker nodes are inline
    while (startNode->get_ParentNode()->get_NodeType() != Aspose::Words::NodeType::Body)
    {
        startNode = startNode->get_ParentNode();
    }

    while (endNode->get_ParentNode()->get_NodeType() != Aspose::Words::NodeType::Body)
    {
        endNode = endNode->get_ParentNode();
    }

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
        System::SharedPtr<CompositeNode> cloneNode = System::DynamicCast<Aspose::Words::CompositeNode>(currNode->Clone(true));
        isEndingNode = currNode->Equals(endNode);

        if (isStartingNode || isEndingNode)
        {
            // We need to process each marker separately so pass it off to a separate method instead.
            if (isStartingNode)
            {
                ProcessMarker(cloneNode, nodes, originalStartNode, isInclusive, isStartingNode, isEndingNode);
                isStartingNode = false;
            }

            // Conditional needs to be separate as the block level start and end markers maybe the same node.
            if (isEndingNode)
            {
                ProcessMarker(cloneNode, nodes, originalEndNode, isInclusive, isStartingNode, isEndingNode);
                isExtracting = false;
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
            System::SharedPtr<Section> nextSection = System::DynamicCast<Aspose::Words::Section>(currNode->GetAncestor(Aspose::Words::NodeType::Section)->get_NextSibling());
            currNode = nextSection->get_Body()->get_FirstChild();
        }
        else
        {
            // Move to the next node in the body.
            currNode = currNode->get_NextSibling();
        }
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
    System::SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true);
    // Look through all paragraphs to find those with the specified style.
    auto paragraph_enumerator = paragraphs->GetEnumerator();
    System::SharedPtr<Paragraph> paragraph;
    while (paragraph_enumerator->MoveNext() && (paragraph = System::DynamicCast<Paragraph>(paragraph_enumerator->get_Current()), true))
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
    System::SharedPtr<NodeImporter> importer = System::MakeObject<NodeImporter>(srcDoc, dstDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting);

    for (auto& node : nodes)
    {
        System::SharedPtr<Node> importNode = importer->ImportNode(node, true);
        dstDoc->get_FirstSection()->get_Body()->AppendChild(importNode);
    }

    // Return the generated document.
    return dstDoc;
}
// ExEnd:CommonGenerateDocument

