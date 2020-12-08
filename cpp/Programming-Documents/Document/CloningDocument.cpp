#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>

using namespace Aspose::Words;

void CloningDocument()
{
    std::cout << "CloningDocument example started." << std::endl;
    // ExStart:CloningDocument
    // Create a Document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"This is the original document before applying the clone method");
    
    // Clone the document.
    System::SharedPtr<Document> clone = doc->Clone();

    // Edit the cloned document.
    builder = System::MakeObject<DocumentBuilder>(clone);
    builder->Write(u"Section 1");
    builder->InsertBreak(BreakType::SectionBreakNewPage);
    builder->Write(u"Section 2");

    // This shows what is in the document originally. The document has two sections.
    std::cout << clone->GetText().Trim() << std::endl << std::endl;

    // Duplicate the last section and append the copy to the end of the document.
    auto lastSectionIdx = clone->get_Sections()->get_Count() - 1;
    System::SharedPtr<Section> newSection = clone->get_Sections()->idx_get(lastSectionIdx)->Clone();
    clone->get_Sections()->Add(newSection);

    // Check what the document contains after we changed it.
    std::cout << clone->GetText().Trim() << std::endl << std::endl;
    // ExEnd:CloningDocument
}
