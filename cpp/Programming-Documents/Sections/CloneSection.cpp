#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>

using namespace Aspose::Words;

void CloneSection()
{
    std::cout << "CloneSection example started." << std::endl;
    // ExStart:CloneSection
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithSections();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
    System::SharedPtr<Section> cloneSection = doc->get_Sections()->idx_get(0)->Clone();
    // ExEnd:CloneSection
    std::cout << "0 index section clone successfully." << std::endl;
    std::cout << "CloneSection example finished." << std::endl << std::endl;
}