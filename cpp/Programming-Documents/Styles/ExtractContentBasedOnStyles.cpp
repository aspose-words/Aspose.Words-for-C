#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include "Model/Document/Document.h"
#include <Model/Document/SaveFormat.h>
#include <Model/Styles/Style.h>
#include <Model/Text/Font.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/Text/Range.h>
#include <Model/Text/Run.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>

using namespace Aspose::Words;

namespace
{
    // ExStart:ParagraphsByStyleName
    std::vector<System::SharedPtr<Paragraph>> ParagraphsByStyleName(System::SharedPtr<Document> doc, System::String const &styleName)
    {
        // Create an array to collect paragraphs of the specified style.
        std::vector<System::SharedPtr<Paragraph>> paragraphsWithStyle;
        // Get all paragraphs from the document.
        System::SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(NodeType::Paragraph, true);
        // Look through all paragraphs to find those with the specified style.
        for (System::SharedPtr<Paragraph> paragraph : System::IterateOver<System::SharedPtr<Paragraph>>(paragraphs))
        {
            if (paragraph->get_ParagraphFormat()->get_Style()->get_Name() == styleName)
            {
                paragraphsWithStyle.push_back(paragraph);
            }
        }
        return paragraphsWithStyle;
    }
    // ExEnd:ParagraphsByStyleName

    // ExStart:RunsByStyleName
    std::vector<System::SharedPtr<Run>> RunsByStyleName(System::SharedPtr<Document> doc, System::String const &styleName)
    {
        // Create an array to collect runs of the specified style.
        std::vector<System::SharedPtr<Run>> runsWithStyle;
        // Get all runs from the document.
        System::SharedPtr<NodeCollection> runs = doc->GetChildNodes(NodeType::Run, true);
        // Look through all runs to find those with the specified style.
        for (System::SharedPtr<Run> run : System::IterateOver<System::SharedPtr<Run>>(runs))
        {
            if (run->get_Font()->get_Style()->get_Name() == styleName)
            {
                runsWithStyle.push_back(run);
            }
        }
        return runsWithStyle;
    }
    // ExEnd:RunsByStyleName
}

void ExtractContentBasedOnStyles()
{
    std::cout << "ExtractContentBasedOnStyles example started." << std::endl;
    // ExStart:ExtractContentBasedOnStyles
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithStyles();
    // Open the document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
    // Define style names as they are specified in the Word document.
    const System::String paraStyle = u"Heading 1";
    const System::String runStyle = u"Intense Emphasis";
    // Collect paragraphs with defined styles.
    // Show the number of collected paragraphs and display the text of this paragraphs.
    std::vector<System::SharedPtr<Paragraph>> paragraphs = ParagraphsByStyleName(doc, paraStyle);
    std::cout << "Paragraphs with \"" << paraStyle.ToUtf8String() << "\" styles (" << paragraphs.size() << "):" << std::endl;
    for (System::SharedPtr<Paragraph> paragraph : paragraphs)
    {
        std::cout << paragraph->ToString(SaveFormat::Text).ToUtf8String();
    }
    std::cout << std::endl;
    // Collect runs with defined styles.
    // Show the number of collected runs and display the text of this runs.
    std::vector<System::SharedPtr<Run>> runs = RunsByStyleName(doc, runStyle);
    std::cout << "Runs with \"" << runStyle.ToUtf8String() << "\" styles (" << runs.size() << "):" << std::endl;
    for (System::SharedPtr<Run> run : runs)
    {
        std::cout << run->get_Range()->get_Text().ToUtf8String() << std::endl;
    }
    // ExEnd:ExtractContentBasedOnStyles
    std::cout << "Extracted contents based on styles successfully." << std::endl;
    std::cout << "ExtractContentBasedOnStyles example finished." << std::endl << std::endl;
}