#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>

using namespace Aspose::Words;

void ChangeStyleOfTOCLevel()
{
    std::cout << "ChangeStyleOfTOCLevel example started." << std::endl;
    // ExStart:ChangeStyleOfTOCLevel
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    // Retrieve the style used for the first level of the TOC and change the formatting of the style.
    doc->get_Styles()->idx_get(StyleIdentifier::Toc1)->get_Font()->set_Bold(true);
    // ExEnd:ChangeStyleOfTOCLevel
    std::cout << "TOC level style changed successfully." << std::endl;
    std::cout << "ChangeStyleOfTOCLevel example finished." << std::endl << std::endl;
}