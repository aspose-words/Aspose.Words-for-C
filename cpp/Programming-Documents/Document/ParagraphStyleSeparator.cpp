#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Model/Document/Document.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Text/Paragraph.h>

using namespace Aspose::Words;

void ParagraphStyleSeparator()
{
    std::cout << "ParagraphStyleSeparator example started." << std::endl;
    // ExStart:ParagraphStyleSeparator
    // The path to the documents directory.
    System::String dataDir = GetDataDir_RenderingAndPrinting();

    // Initialize document.
    System::String fileName = u"TestFile.doc";
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + fileName);

    for (auto paragraph : System::IterateOver<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)))
    {
        if (paragraph->get_BreakIsStyleSeparator())
        {
            std::cout << "Separator Found!" << std::endl;
        }
    }
    // ExEnd:ParagraphStyleSeparator
    std::cout << "ParagraphStyleSeparator example finished." << std::endl << std::endl;
}