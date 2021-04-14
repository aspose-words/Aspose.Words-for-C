#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/TabAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/TabLeader.h>
#include <Aspose.Words.Cpp/Model/Text/TabStop.h>
#include <Aspose.Words.Cpp/Model/Text/TabStopCollection.h>
#include <system/enumerator_adapter.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Managment {

class WorkingWithTableOfContent : public DocsExamplesBase
{
public:
    void ChangeStyleOfTocLevel()
    {
        //ExStart:ChangeStyleOfTOCLevel
        auto doc = MakeObject<Document>();
        // Retrieve the style used for the first level of the TOC and change the formatting of the style.
        doc->get_Styles()->idx_get(StyleIdentifier::Toc1)->get_Font()->set_Bold(true);
        //ExEnd:ChangeStyleOfTOCLevel
    }

    void ChangeTocTabStops()
    {
        //ExStart:ChangeTOCTabStops
        auto doc = MakeObject<Document>(MyDir + u"Table of contents.docx");

        for (auto para : System::IterateOver<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)))
        {
            // Check if this paragraph is formatted using the TOC result based styles.
            // This is any style between TOC and TOC9.
            if (para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() >= StyleIdentifier::Toc1 &&
                para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() <= StyleIdentifier::Toc9)
            {
                // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                SharedPtr<TabStop> tab = para->get_ParagraphFormat()->get_TabStops()->idx_get(0);

                // Remove the old tab from the collection.
                para->get_ParagraphFormat()->get_TabStops()->RemoveByPosition(tab->get_Position());

                // Insert a new tab using the same properties but at a modified position.
                // We could also change the separators used (dots) by passing a different Leader type.
                para->get_ParagraphFormat()->get_TabStops()->Add(tab->get_Position() - 50, tab->get_Alignment(), tab->get_Leader());
            }
        }

        doc->Save(ArtifactsDir + u"WorkingWithTableOfContent.ChangeTocTabStops.docx");
        //ExEnd:ChangeTOCTabStops
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Contents_Managment
