#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>

using namespace Aspose::Words;

void AppendSectionContent()
{
    std::cout << "AppendSectionContent example started." << std::endl;
    // ExStart:AppendSectionContent
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithSections();
    auto doc = System::MakeObject<Document>(dataDir + u"Section.AppendContent.doc");
    // This is the section that we will append and prepend to.
    auto section = doc->get_Sections()->idx_get(2);
    // This copies content of the 1st section and inserts it at the beginning of the specified section.
    auto sectionToPrepend = doc->get_Sections()->idx_get(0);
    section->PrependContent(sectionToPrepend);
    // This copies content of the 2nd section and inserts it at the end of the specified section.
    auto sectionToAppend = doc->get_Sections()->idx_get(1);
    section->AppendContent(sectionToAppend);
    // ExEnd:AppendSectionContent
    std::cout << "Section content appended successfully." << std::endl;
    std::cout << "AppendSectionContent example finished." << std::endl << std::endl;
}