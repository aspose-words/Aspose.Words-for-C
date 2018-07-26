#include "../../examples.h"

#include "system/shared_ptr.h"

#include "Model/Document/Document.h"
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>

using namespace Aspose::Words;

void DeleteSectionContent()
{
	// ExStart:DeleteSectionContent
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithSections();
    auto doc = System::MakeObject<Document>(dataDir + u"Document.doc");
    auto section = doc->get_Sections()->idx_get(0);
    section->ClearContent();
    // ExEnd:DeleteSectionContent
    std::cout << "\nSection content at 0 index deleted successfully." << std::endl;
}