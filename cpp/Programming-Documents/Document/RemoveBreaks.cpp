#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

namespace
{
    // ExStart:RemovePageBreaks
    void RemovePageBreaks(const System::SharedPtr<Document>& doc)
    {
        // Retrieve all paragraphs in the document.
        System::SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(NodeType::Paragraph, true);

        // Iterate through all paragraphs
        for (System::SharedPtr<Paragraph> para : System::IterateOver<System::SharedPtr<Paragraph>>(paragraphs))
        {
            // If the paragraph has a page break before set then clear it.
            if (para->get_ParagraphFormat()->get_PageBreakBefore())
            {
                para->get_ParagraphFormat()->set_PageBreakBefore(false);
            }
            // Check all runs in the paragraph for page breaks and remove them.
            for (System::SharedPtr<Run> run : System::IterateOver<System::SharedPtr<Run>>(para->get_Runs()))
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
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    // ExStart:OpenFromFile
    // Open the document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
    // ExEnd:OpenFromFile

    // Remove the page and section breaks from the document.
    // In Aspose.Words section breaks are represented as separate Section nodes in the document.
    // To remove these separate sections the sections are combined.
    RemovePageBreaks(doc);
    RemoveSectionBreaks(doc);

    System::String outputPath = outputDataDir + u"RemoveBreaks.doc";
    // Save the document.
    doc->Save(outputPath);
    std::cout << "Removed breaks from the document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "RemoveBreaks example finished." << std::endl << std::endl;
}
