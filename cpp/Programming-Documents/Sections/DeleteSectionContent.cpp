#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>

using namespace Aspose::Words;

void DeleteSectionContent()
{
    std::cout << "DeleteSectionContent example started." << std::endl;
    // ExStart:DeleteSectionContent
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithSections();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
    System::SharedPtr<Section> section = doc->get_Sections()->idx_get(0);
    section->ClearContent();
    // ExEnd:DeleteSectionContent
    std::cout << "Section content at 0 index deleted successfully." << std::endl;
    std::cout << "DeleteSectionContent example finished." << std::endl << std::endl;
}