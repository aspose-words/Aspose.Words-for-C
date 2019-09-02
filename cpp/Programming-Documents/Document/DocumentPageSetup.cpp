#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/SectionLayoutMode.h>
#include <Model/Sections/PageSetup.h>

using namespace Aspose::Words;

void DocumentPageSetup()
{
    std::cout << "DocumentPageSetup example started." << std::endl;
    // ExStart:DocumentPageSetup
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

    //Set the layout mode for a section allowing to define the document grid behavior
    //Note that the Document Grid tab becomes visible in the Page Setup dialog of MS Word if any Asian language is defined as editing language.
    doc->get_FirstSection()->get_PageSetup()->set_LayoutMode(SectionLayoutMode::Grid);
    //Set the number of characters per line in the document grid. 
    doc->get_FirstSection()->get_PageSetup()->set_CharactersPerLine(30);
    //Set the number of lines per page in the document grid. 
    doc->get_FirstSection()->get_PageSetup()->set_LinesPerPage(10);

    System::String outputPath = outputDataDir + u"DocumentPageSetup.doc";
    doc->Save(outputPath);
    // ExEnd:DocumentPageSetup
    std::cout << "PageSetup properties are set successfully." << std::endl;
    std::cout << "DocumentPageSetup example finished." << std::endl << std::endl;
}
