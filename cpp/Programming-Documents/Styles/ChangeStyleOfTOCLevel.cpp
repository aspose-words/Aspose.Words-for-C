#include "stdafx.h"
#include "examples.h"

#include <system/shared_ptr.h>

#include "Model/Document/Document.h"
#include <Model/Styles/StyleIdentifier.h>
#include <Model/Styles/StyleCollection.h>
#include <Model/Text/Font.h>

using namespace Aspose::Words;

void ChangeStyleOfTOCLevel()
{
    // ExStart:ChangeStyleOfTOCLevel
    auto doc = System::MakeObject<Document>();
    // Retrieve the style used for the first level of the TOC and change the formatting of the style.
    doc->get_Styles()->idx_get(Aspose::Words::StyleIdentifier::Toc1)->get_Font()->set_Bold(true);
    // ExEnd:ChangeStyleOfTOCLevel
    std::cout << "\nTOC level style changed successfully." << std::endl;
}