#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>

using namespace Aspose::Words;

void ParagraphStyleSeparator()
{
    std::cout << "ParagraphStyleSeparator example started." << std::endl;
    // ExStart:ParagraphStyleSeparator
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();

    // Initialize document.
    System::String fileName = u"TestFile.doc";
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + fileName);

    for (System::SharedPtr<Paragraph> paragraph : System::IterateOver<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)))
    {
        if (paragraph->get_BreakIsStyleSeparator())
        {
            std::cout << "Separator Found!" << std::endl;
        }
    }
    // ExEnd:ParagraphStyleSeparator
    std::cout << "ParagraphStyleSeparator example finished." << std::endl << std::endl;
}