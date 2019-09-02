#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>

using namespace Aspose::Words;

void CopySection()
{
    std::cout << "CopySection example started." << std::endl;
    // ExStart:CopySection
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithSections();
    System::String outputDataDir = GetOutputDataDir_WorkingWithSections();
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>();
    System::SharedPtr<Section> sourceSection = srcDoc->get_Sections()->idx_get(0);
    System::SharedPtr<Section> newSection = System::DynamicCast<Section>(dstDoc->ImportNode(sourceSection, true));
    dstDoc->get_Sections()->Add(newSection);
    System::String outputPath = outputDataDir + u"CopySection.doc";
    dstDoc->Save(outputPath);
    // ExEnd:CopySection
    std::cout << "Section copied successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CopySection example finished." << std::endl << std::endl;
}