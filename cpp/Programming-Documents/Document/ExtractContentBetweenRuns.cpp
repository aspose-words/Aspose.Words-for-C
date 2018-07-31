#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/collections/list.h>
#include <system/collections/ilist.h>
#include <Model/Text/RunCollection.h>
#include <Model/Text/Run.h>
#include <Model/Text/Paragraph.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/Node.h>
#include <Model/Document/Document.h>

#include "Common.h"

using namespace Aspose::Words;

void ExtractContentBetweenRuns()
{
    // ExStart:ExtractContentBetweenRuns
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    
    // Retrieve a paragraph from the first section.
    System::SharedPtr<Paragraph> para = System::DynamicCast<Aspose::Words::Paragraph>(doc->GetChild(Aspose::Words::NodeType::Paragraph, 7, true));
    
    // Use some runs for extraction.
    System::SharedPtr<Aspose::Words::Run> startRun = para->get_Runs()->idx_get(1);
    System::SharedPtr<Aspose::Words::Run> endRun = para->get_Runs()->idx_get(4);
    
    // Extract the content between these nodes in the document. Include these markers in the extraction.
    auto extractedNodes = ExtractContent(startRun, endRun, true);
    
    // Get the node from the list. There should only be one paragraph returned in the list.
    System::SharedPtr<Node> node = extractedNodes[0];
    // Print the text of this node to the console.
    std::cout << node->GetText().ToUtf8String() << '\n';
    // ExEnd:ExtractContentBetweenRuns
    std::cout << "\nExtracted content betweenn the runs successfully.\n";
}
