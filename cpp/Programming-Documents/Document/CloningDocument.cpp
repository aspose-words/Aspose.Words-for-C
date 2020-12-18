#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>

using namespace Aspose::Words;

namespace
{
    void DeepDocumentCopy(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:CloningDocument
        auto doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
        auto clone = doc->Clone();

        // Save the document to disk.
        clone->Save(outputDataDir + u"TestFile.DeepDocumentCopy.doc");
    }

    void DuplicateSection(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:DuplicateSection
        // Create a document.
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"This is the original document before applying the clone method");

        // Clone the document.
        auto clone = doc->Clone();

        // Edit the cloned document.
        builder = System::MakeObject<DocumentBuilder>(clone);
        builder->Write(u"Section 1");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Write(u"Section 2");

        // This shows what is in the document originally. The document has two sections.
        std::cout << clone->GetText().Trim() << std::endl << std::endl;

        // Duplicate the last section and append the copy to the end of the document.
        auto lastSectionIdx = clone->get_Sections()->get_Count() - 1;
        auto newSection = clone->get_Sections()->idx_get(lastSectionIdx)->Clone();
        clone->get_Sections()->Add(newSection);
        // ExEnd:DuplicateSection

        // Check what the document contains after we changed it.
        std::cout << clone->GetText().Trim() << std::endl << std::endl;
    }

}

void CloningDocument()
{
    auto inputDataDir = GetInputDataDir_WorkingWithDocument();
    auto outputDataDir = GetOutputDataDir_WorkingWithDocument();

    std::cout << "CloningDocument example started." << std::endl;
}
