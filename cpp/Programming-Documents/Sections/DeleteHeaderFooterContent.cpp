#include "../../examples.h"

#include "system/shared_ptr.h"

#include "Model/Document/Document.h"
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>

using namespace Aspose::Words;

void DeleteHeaderFooterContent()
{
    // ExStart:DeleteHeaderFooterContent
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithSections();
    auto doc = System::MakeObject<Document>(dataDir + u"Document.doc");
    auto section = doc->get_Sections()->idx_get(0);
    section->ClearHeadersFooters();
    // ExEnd:DeleteHeaderFooterContent
    std::cout << "\nHeader and footer content of 0 index deleted successfully." << std::endl;
}