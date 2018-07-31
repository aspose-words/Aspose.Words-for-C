#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/collections/ienumerator.h>
#include <Model/Styles/StyleCollection.h>
#include <Model/Styles/Style.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void AccessStyles()
{
    // ExStart:AccessStyles
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Load the template document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    // Get styles collection from document.
    System::SharedPtr<StyleCollection> styles = doc->get_Styles();
    System::String styleName = u"";
    // Iterate through all the styles.
    auto style_enumerator = styles->GetEnumerator();
    System::SharedPtr<Style> style;
    while (style_enumerator->MoveNext() && (style = style_enumerator->get_Current(), true))
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
    std::cout << "\nDocument have following styles " << styleName.ToUtf8String() << '\n';
}
