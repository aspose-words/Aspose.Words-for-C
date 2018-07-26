#include "../../examples.h"

#include "system/shared_ptr.h"

#include "Model/Document/Document.h"
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>

using namespace Aspose::Words;

void CopySection()
{
    // ExStart:CopySection
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithSections();
    auto srcDoc = System::MakeObject<Document>(dataDir + u"Document.doc");
    auto dstDoc = System::MakeObject<Document>();
    auto sourceSection = srcDoc->get_Sections()->idx_get(0);
    auto newSection = System::DynamicCast<Aspose::Words::Section>(dstDoc->ImportNode(sourceSection, true));
    dstDoc->get_Sections()->Add(newSection);
    System::String outputPath = dataDir + GetOutputFilePath(u"CopySection.doc");
    dstDoc->Save(outputPath);
    // ExEnd:CopySection
    std::cout << "\nSection copied successfully.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
}