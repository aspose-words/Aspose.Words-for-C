#include "stdafx.h"
#include "examples.h"

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
    // ExStart:ChangeTOCTabStops
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithStyles();
    // Open the document.
    auto doc = System::MakeObject<Document>(dataDir + u"Document.TableOfContents.doc");
    // Iterate through all paragraphs in the document
    auto para_enumerator = (System::DynamicCastEnumerableTo<System::SharedPtr<Paragraph>>(doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)))->GetEnumerator();
    decltype(para_enumerator->get_Current()) para;
    while (para_enumerator->MoveNext() && (para = para_enumerator->get_Current(), true))
    {
        // Check if this paragraph is formatted using the TOC result based styles. This is any style between TOC and TOC9.
        if (para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() >= Aspose::Words::StyleIdentifier::Toc1 && para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() <= Aspose::Words::StyleIdentifier::Toc9)
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
    std::cout << "\nPosition of the right tab stop in TOC related paragraphs modified successfully.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
}