#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Model/Styles/StyleCollection.h>
#include <Model/Styles/Style.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void AccessStyles()
{
    std::cout << "AccessStyles example started." << std::endl;
    // ExStart:AccessStyles
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    // Load the template document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
    // Get styles collection from document.
    System::SharedPtr<StyleCollection> styles = doc->get_Styles();
    System::String styleName = u"";
    // Iterate through all the styles.
    for (System::SharedPtr<Style> style : System::IterateOver(styles))
    {
        if (styleName == u"")
        {
            styleName = style->get_Name();
        }
        else
        {
            styleName = styleName + u", " + style->get_Name();
        }
    }
    // ExEnd:AccessStyles
    std::cout << "Document have following styles " << styleName.ToUtf8String() << std::endl;
    std::cout << "AccessStyles example finished." << std::endl << std::endl;
}
