#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Layout/Public/LayoutCollector.h>
#include <Model/Document/Document.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/Section.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;

void PageNumbersOfNodes()
{
    std::cout << "PageNumbersOfNodes example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.docx");

    // Create and attach collector before the document before page layout is built.
    System::SharedPtr<LayoutCollector> layoutCollector = System::MakeObject<LayoutCollector>(doc);

    // This will build layout model and collect necessary information.
    doc->UpdatePageLayout();

    // Print the details of each document node including the page numbers. 
    for (System::SharedPtr<Node> node : System::IterateOver(doc->get_FirstSection()->get_Body()->GetChildNodes(NodeType::Any, true)))
    {
        std::cout << " --------- " << std::endl;
        std::cout << "NodeType:   " << Node::NodeTypeToString(node->get_NodeType()).ToUtf8String() << std::endl;
        std::cout << "Text:       \"" << node->ToString(SaveFormat::Text).Trim().ToUtf8String() << "\"" << std::endl;
        std::cout << "Page Start: " << layoutCollector->GetStartPageIndex(node) << std::endl;
        std::cout << "Page End:   " << layoutCollector->GetEndPageIndex(node) << std::endl;
        std::cout << " --------- " << std::endl << std::endl;
    }

    // Detatch the collector from the document.
    layoutCollector->set_Document(nullptr);
    std::cout << "Found the page numbers of all nodes successfully." << std::endl;
    std::cout << "PageNumbersOfNodes example finished." << std::endl << std::endl;
}