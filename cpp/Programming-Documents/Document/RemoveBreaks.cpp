#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/collections/ienumerator.h>
#include <Model/Text/RunCollection.h>
#include <Model/Text/Run.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/ControlChar.h>
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Document/Document.h>
#include <cstdint>

using namespace Aspose::Words;

namespace
{
    // ExStart:RemovePageBreaks
    void RemovePageBreaks(const System::SharedPtr<Document>& doc)
    {
        // Retrieve all paragraphs in the document.
        System::SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(NodeType::Paragraph, true);

        // Iterate through all paragraphs
        for (System::SharedPtr<Paragraph> para : System::IterateOver(System::DynamicCastEnumerableTo<System::SharedPtr<Paragraph>>(paragraphs)))
        {
            // If the paragraph has a page break before set then clear it.
            if (para->get_ParagraphFormat()->get_PageBreakBefore())
            {
                para->get_ParagraphFormat()->set_PageBreakBefore(false);
            }
            // Check all runs in the paragraph for page breaks and remove them.
            for (System::SharedPtr<Run> run : System::IterateOver(System::DynamicCastEnumerableTo<System::SharedPtr<Run>>(para->get_Runs())))
            {
                if (run->get_Text().Contains(ControlChar::PageBreak()))
                {
                    run->set_Text(run->get_Text().Replace(ControlChar::PageBreak(), System::String::Empty));
                }
            }
        }
    }
    // ExEnd:RemovePageBreaks

    // ExStart:RemoveSectionBreaks
    void RemoveSectionBreaks(const System::SharedPtr<Document>& doc)
    {
        // Loop through all sections starting from the section that precedes the last one 
        // And moving to the first section.
        for (int32_t i = doc->get_Sections()->get_Count() - 2; i >= 0; i--)
        {
            // Copy the content of the current section to the beginning of the last section.
            doc->get_LastSection()->PrependContent(doc->get_Sections()->idx_get(i));
            // Remove the copied section.
            doc->get_Sections()->idx_get(i)->Remove();
        }
    }
    // ExEnd:RemoveSectionBreaks
}

void RemoveBreaks()
{
    std::cout << "RemoveBreaks example started." << std::endl;
    // ExStart:RemoveBreaks
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // ExStart:OpenFromFile
    // Open the document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    // ExEnd:OpenFromFile

    // Remove the page and section breaks from the document.
    // In Aspose.Words section breaks are represented as separate Section nodes in the document.
    // To remove these separate sections the sections are combined.
    RemovePageBreaks(doc);
    RemoveSectionBreaks(doc);

    System::String outputPath = dataDir + GetOutputFilePath(u"RemoveBreaks.doc");
    // Save the document.
    doc->Save(outputPath);
    // ExEnd:RemoveBreaks
    std::cout << "Removed breaks from the document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "RemoveBreaks example finished." << std::endl << std::endl;
}
