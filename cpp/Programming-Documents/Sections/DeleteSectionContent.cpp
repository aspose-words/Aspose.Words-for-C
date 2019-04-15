#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>

using namespace Aspose::Words;

void DeleteSectionContent()
{
    std::cout << "DeleteSectionContent example started." << std::endl;
    // ExStart:DeleteSectionContent
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithSections();
    auto doc = System::MakeObject<Document>(dataDir + u"Document.doc");
    auto section = doc->get_Sections()->idx_get(0);
    section->ClearContent();
    // ExEnd:DeleteSectionContent
    std::cout << "Section content at 0 index deleted successfully." << std::endl;
    std::cout << "DeleteSectionContent example finished." << std::endl << std::endl;
}