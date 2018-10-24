#include "stdafx.h"
#include "examples.h"

#include <system/shared_ptr.h>
#include <system/special_casts.h>
#include <system/collections/list.h>

#include "Model/Document/Document.h"
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
    typedef System::Collections::Generic::List<System::SharedPtr<Paragraph>> TParagraphList;
    typedef System::SharedPtr<TParagraphList> TParagraphListP;

    TParagraphListP ParagraphsByStyleName(System::SharedPtr<Document> doc, System::String const &styleName)
    {
        // Create an array to collect paragraphs of the specified style.
        auto paragraphsWithStyle = System::MakeObject<TParagraphList>();
        // Get all paragraphs from the document.
        auto paragraphs = doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true);
        // Look through all paragraphs to find those with the specified style.
        auto paragraph_enumerator = (System::DynamicCastEnumerableTo<System::SharedPtr<Paragraph>>(paragraphs))->GetEnumerator();
        decltype(paragraph_enumerator->get_Current()) paragraph;
        while (paragraph_enumerator->MoveNext() && (paragraph = paragraph_enumerator->get_Current(), true))
        {
            if (paragraph->get_ParagraphFormat()->get_Style()->get_Name() == styleName)
            {
                paragraphsWithStyle->Add(paragraph);
            }
        }
        return paragraphsWithStyle;
    }

    typedef System::Collections::Generic::List<System::SharedPtr<Run>> TRunList;
    typedef System::SharedPtr<TRunList> TRunListP;

    TRunListP RunsByStyleName(System::SharedPtr<Document> doc, System::String const &styleName)
    {
        // Create an array to collect runs of the specified style.
        auto runsWithStyle = System::MakeObject<TRunList>();
        // Get all runs from the document.
        auto runs = doc->GetChildNodes(Aspose::Words::NodeType::Run, true);
        // Look through all runs to find those with the specified style.
        auto run_enumerator = (System::DynamicCastEnumerableTo<System::SharedPtr<Aspose::Words::Run>>(runs))->GetEnumerator();
        decltype(run_enumerator->get_Current()) run;
        while (run_enumerator->MoveNext() && (run = run_enumerator->get_Current(), true))
        {
            if (run->get_Font()->get_Style()->get_Name() == styleName)
            {
                runsWithStyle->Add(run);
            }
        }
        return runsWithStyle;
    }
}

void ExtractContentBasedOnStyles()
{
    std::cout << "ExtractContentBasedOnStyles example started." << std::endl;
    // ExStart:ExtractContentBasedOnStyles
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithStyles();
    // Open the document.
    auto doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    // Define style names as they are specified in the Word document.
    const System::String paraStyle = u"Heading 1";
    const System::String runStyle = u"Intense Emphasis";
    // Collect paragraphs with defined styles.
    // Show the number of collected paragraphs and display the text of this paragraphs.
    auto paragraphs = ParagraphsByStyleName(doc, paraStyle);
    std::cout << "Paragraphs with \"" << paraStyle.ToUtf8String() << "\" styles (" << paragraphs->get_Count() << "):" << std::endl;
    auto paragraph_enumerator = (paragraphs)->GetEnumerator();
    decltype(paragraph_enumerator->get_Current()) paragraph;
    while (paragraph_enumerator->MoveNext() && (paragraph = paragraph_enumerator->get_Current(), true))
    {
        std::cout << paragraph->GetText().ToUtf8String();
    }
    std::cout << std::endl;
    // Collect runs with defined styles.
    // Show the number of collected runs and display the text of this runs.
    auto runs = RunsByStyleName(doc, runStyle);
    std::cout << "Runs with \"" << runStyle.ToUtf8String() << "\" styles (" << runs->get_Count() << "):" << std::endl;
    auto run_enumerator = (runs)->GetEnumerator();
    decltype(run_enumerator->get_Current()) run;
    while (run_enumerator->MoveNext() && (run = run_enumerator->get_Current(), true))
    {
        std::cout << run->get_Range()->get_Text().ToUtf8String() << std::endl;
    }
    // ExEnd:ExtractContentBasedOnStyles
    std::cout << "Extracted contents based on styles successfully." << std::endl;
    std::cout << "ExtractContentBasedOnStyles example finished." << std::endl << std::endl;
}