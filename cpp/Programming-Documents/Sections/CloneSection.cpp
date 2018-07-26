#include "../../examples.h"

#include "system/shared_ptr.h"

#include "Model/Document/Document.h"
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>

using namespace Aspose::Words;

void CloneSection()
{
    // ExStart:CloneSection
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithSections();
    auto doc = System::MakeObject<Document>(dataDir + u"Document.doc");
    auto cloneSection = doc->get_Sections()->idx_get(0)->Clone();
    // ExEnd:CloneSection
    std::cout << "\n0 index section clone successfully." << std::endl;
}