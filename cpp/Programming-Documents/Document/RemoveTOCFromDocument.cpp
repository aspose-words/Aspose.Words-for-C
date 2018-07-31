#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/collections/list.h>
#include <system/collections/ilist.h>
#include <system/collections/ienumerator.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/Node.h>
#include <Model/Fields/Nodes/FieldStart.h>
#include <Model/Fields/Nodes/FieldEnd.h>
#include <Model/Fields/FieldType.h>
#include <Model/Document/Document.h>
#include <cstdint>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace
{

// ExStart:RemoveTOCFromDocument
/// <summary>
/// Removes the specified table of contents field from the document.
/// </summary>
/// <param name="doc">The document to remove the field from.</param>
/// <param name="index">The zero-based index of the TOC to remove.</param>
void RemoveTableOfContents(const System::SharedPtr<Document>& doc, int32_t index)
{
    // Store the FieldStart nodes of TOC fields in the document for quick access.
    std::vector<System::SharedPtr<FieldStart>> fieldStarts;
    // This is a list to store the nodes found inside the specified TOC. They will be removed
    // At the end of this method.
    std::vector<System::SharedPtr<Node>> nodeList;
    
    auto start_enumerator = doc->GetChildNodes(Aspose::Words::NodeType::FieldStart, true)->GetEnumerator();
    System::SharedPtr<FieldStart> start;
    while (start_enumerator->MoveNext() && (start = System::DynamicCast<FieldStart>(start_enumerator->get_Current()), true))
    {
        if (start->get_FieldType() == Aspose::Words::Fields::FieldType::FieldTOC)
        {
            // Add all FieldStarts which are of type FieldTOC.
            fieldStarts.push_back(start);
        }
    }
    
    // Ensure the TOC specified by the passed index exists.
    if (index > fieldStarts.size() - 1)
    {
        throw System::ArgumentOutOfRangeException(u"TOC index is out of range");
    }
    
    bool isRemoving = true;
    // Get the FieldStart of the specified TOC.
    System::SharedPtr<Node> currentNode = fieldStarts[index];
    
    while (isRemoving)
    {
        // It is safer to store these nodes and delete them all at once later.
        nodeList.push_back(currentNode);
        currentNode = currentNode->NextPreOrder(doc);
        
        // Once we encounter a FieldEnd node of type FieldTOC then we know we are at the end
        // Of the current TOC and we can stop here.
        if (currentNode->get_NodeType() == Aspose::Words::NodeType::FieldEnd)
        {
            System::SharedPtr<FieldEnd> fieldEnd = System::DynamicCast<Aspose::Words::Fields::FieldEnd>(currentNode);
            if (fieldEnd->get_FieldType() == Aspose::Words::Fields::FieldType::FieldTOC)
            {
                isRemoving = false;
            }
        }
    }
    
    // Remove all nodes found in the specified TOC.

    for (auto& node : nodeList)
    {
        node->Remove();
    }
}
// ExEnd:RemoveTOCFromDocument

}

void RemoveTOCFromDocument()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithStyles();
    
    // Open a document which contains a TOC.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.TableOfContents.doc");
    
    // Remove the first table of contents from the document.
    RemoveTableOfContents(doc, 0);
    
    dataDir = dataDir + GetOutputFilePath(u"RemoveTOCFromDocument.doc");
    // Save the output.
    doc->Save(dataDir);
    
    std::cout << "\nSpecified TOC from a document removed successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
