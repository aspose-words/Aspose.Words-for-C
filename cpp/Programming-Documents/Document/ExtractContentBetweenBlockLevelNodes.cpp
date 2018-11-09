#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/console.h>
#include <system/collections/list.h>
#include <Model/Text/Paragraph.h>
#include <Model/Tables/Table.h>
#include <Model/Sections/Section.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/CompositeNode.h>
#include <Model/Document/Document.h>
#include <cstdint>

#include "Common.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

void ExtractContentBetweenBlockLevelNodes()
{
    std::cout << "ExtractContentBetweenBlockLevelNodes example started." << std::endl;
    // ExStart:ExtractContentBetweenBlockLevelNodes
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");

    System::SharedPtr<Paragraph> startPara = System::DynamicCast<Paragraph>(doc->get_LastSection()->GetChild(NodeType::Paragraph, 2, true));
    System::SharedPtr<Table> endTable = System::DynamicCast<Table>(doc->get_LastSection()->GetChild(NodeType::Table, 0, true));

    // Extract the content between these nodes in the document. Include these markers in the extraction.
    auto extractedNodes = ExtractContent(startPara, endTable, true);

    // Lets reverse the array to make inserting the content back into the document easier.
    for (auto it = extractedNodes.rbegin(); it != extractedNodes.rend(); ++it)
    {
        endTable->get_ParentNode()->InsertAfter(*it, endTable);
    }

    System::String outputPath = dataDir + GetOutputFilePath(u"ExtractContentBetweenBlockLevelNodes.doc");
    // Save the generated document to disk.
    doc->Save(outputPath);
    // ExEnd:ExtractContentBetweenBlockLevelNodes
    std::cout << "Extracted content betweenn the block level nodes successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ExtractContentBetweenBlockLevelNodes example finished." << std::endl << std::endl;
}
