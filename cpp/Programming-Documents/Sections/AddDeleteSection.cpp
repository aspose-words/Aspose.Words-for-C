#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>

using namespace Aspose::Words;

namespace
{
    void AddSection(System::String const &documentPath)
    {
        // ExStart:AddSection
        auto doc = System::MakeObject<Document>(documentPath);
        auto sectionToAdd = System::MakeObject<Section>(doc);
        doc->get_Sections()->Add(sectionToAdd);
        // ExEnd:AddSection
        std::cout << "Section added successfully to the end of the document." << std::endl;
    }

    void DeleteSection(System::String const &documentPath)
    {
        // ExStart:DeleteSection
        auto doc = System::MakeObject<Document>(documentPath);
        doc->get_Sections()->RemoveAt(0);
        // ExEnd:DeleteSection
        std::cout << "Section deleted successfully at 0 index." << std::endl;
    }

    void DeleteAllSections(System::String const &documentPath)
    {
        // ExStart:DeleteAllSections
        auto doc = System::MakeObject<Document>(documentPath);
        doc->get_Sections()->Clear();
        // ExEnd:DeleteAllSections
        std::cout << "All sections deleted successfully form the document." << std::endl;
    }
}

void AddDeleteSection()
{
    std::cout << "AddDeleteSection example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithSections();
    // The path to the document
    System::String documentPath = dataDir + u"Section.AddRemove.doc";
    AddSection(documentPath);
    DeleteSection(documentPath);
    DeleteAllSections(documentPath);
    std::cout << "AddDeleteSection example finished." << std::endl << std::endl;
}