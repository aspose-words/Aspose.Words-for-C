#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void RemoveFooters()
{
    std::cout << "RemoveFooters example started." << std::endl;
    // ExStart:RemoveFooters
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"HeaderFooter.RemoveFooters.doc");

    for (System::SharedPtr<Section> section : System::IterateOver<System::SharedPtr<Section>>(doc))
    {
        // Up to three different footers are possible in a section (for first, even and odd pages).
        // We check and delete all of them.
        System::SharedPtr<HeaderFooter> footer;

        footer = section->get_HeadersFooters()->idx_get(HeaderFooterType::FooterFirst);
        if (footer != nullptr)
        {
            footer->Remove();
        }

        // Primary footer is the footer used for odd pages.
        footer = section->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary);
        if (footer != nullptr)
        {
            footer->Remove();
        }

        footer = section->get_HeadersFooters()->idx_get(HeaderFooterType::FooterEven);
        if (footer != nullptr)
        {
            footer->Remove();
        }
    }

    System::String outputPath = outputDataDir + u"RemoveFooters.doc";
    // Save the document.
    doc->Save(outputPath);
    // ExEnd:RemoveFooters
    std::cout << "All footers from all sections deleted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "RemoveFooters example finished." << std::endl << std::endl;
}
