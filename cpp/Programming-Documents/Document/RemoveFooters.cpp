#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/collections/ienumerator.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/HeaderFooterType.h>
#include <Model/Sections/HeaderFooterCollection.h>
#include <Model/Sections/HeaderFooter.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void RemoveFooters()
{
    std::cout << "RemoveFooters example started." << std::endl;
    // ExStart:RemoveFooters
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"HeaderFooter.RemoveFooters.doc");

    for (System::SharedPtr<Section> section : System::IterateOver(System::DynamicCastEnumerableTo<System::SharedPtr<Section>>(doc)))
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

    System::String outputPath = dataDir + GetOutputFilePath(u"RemoveFooters.doc");
    // Save the document.
    doc->Save(outputPath);
    // ExEnd:RemoveFooters
    std::cout << "All footers from all sections deleted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "RemoveFooters example finished." << std::endl << std::endl;
}
