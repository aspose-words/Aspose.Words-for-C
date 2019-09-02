#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>

using namespace Aspose::Words;

void DeleteHeaderFooterContent()
{
    std::cout << "DeleteHeaderFooterContent example started." << std::endl;
    // ExStart:DeleteHeaderFooterContent
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithSections();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
    System::SharedPtr<Section> section = doc->get_Sections()->idx_get(0);
    section->ClearHeadersFooters();
    // ExEnd:DeleteHeaderFooterContent
    std::cout << "Header and footer content of 0 index deleted successfully." << std::endl;
    std::cout << "DeleteHeaderFooterContent example finished." << std::endl << std::endl;
}