#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <system/shared_ptr.h>
#include <system/special_casts.h>

#include "Model/Document/Document.h"
#include <Model/Styles/Style.h>
#include <Model/Styles/StyleIdentifier.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/Text/TabStop.h>
#include <Model/Text/TabStopCollection.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>

using namespace Aspose::Words;

void ChangeTOCTabStops()
{
    std::cout << "ChangeTOCTabStops example started." << std::endl;
    // ExStart:ChangeTOCTabStops
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithStyles();
    // Open the document.
    auto doc = System::MakeObject<Document>(dataDir + u"Document.TableOfContents.doc");
    // Iterate through all paragraphs in the document
    for (System::SharedPtr<Paragraph> para : System::IterateOver<System::SharedPtr<Paragraph>>(doc->GetChildNodes(NodeType::Paragraph, true)))
    {
        // Check if this paragraph is formatted using the TOC result based styles. This is any style between TOC and TOC9.
        if (para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() >= StyleIdentifier::Toc1 && para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() <= StyleIdentifier::Toc9)
        {
            // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
            auto tab = para->get_ParagraphFormat()->get_TabStops()->idx_get(0);
            // Remove the old tab from the collection.
            para->get_ParagraphFormat()->get_TabStops()->RemoveByPosition(tab->get_Position());
            // Insert a new tab using the same properties but at a modified position.
            // We could also change the separators used (dots) by passing a different Leader type
            para->get_ParagraphFormat()->get_TabStops()->Add(tab->get_Position() - 50, tab->get_Alignment(), tab->get_Leader());
        }
    }
    System::String outputPath = dataDir + GetOutputFilePath(u"ChangeTOCTabStops.doc");
    doc->Save(outputPath);
    // ExEnd:ChangeTOCTabStops
    std::cout << "Position of the right tab stop in TOC related paragraphs modified successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ChangeTOCTabStops example finished." << std::endl << std::endl;
}