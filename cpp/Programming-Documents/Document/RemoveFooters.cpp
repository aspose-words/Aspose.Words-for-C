#include "stdafx.h"
#include "examples.h"

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
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void RemoveFooters()
{
    // ExStart:RemoveFooters
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"HeaderFooter.RemoveFooters.doc");
    
    auto section_enumerator = doc->GetEnumerator();
    System::SharedPtr<Section> section;
    while (section_enumerator->MoveNext() && (section = System::DynamicCast<Section>(section_enumerator->get_Current()), true))
    {
        // Up to three different footers are possible in a section (for first, even and odd pages).
        // We check and delete all of them.
        System::SharedPtr<HeaderFooter> footer;

        footer = section->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterFirst);
        if (footer != nullptr)
        {
            footer->Remove();
        }

        // Primary footer is the footer used for odd pages.
        footer = section->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary);
        if (footer != nullptr)
        {
            footer->Remove();
        }

        footer = section->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterEven);
        if (footer != nullptr)
        {
            footer->Remove();
        }
    }
    dataDir = dataDir + GetOutputFilePath(u"RemoveFooters.doc");
    
    // Save the document.
    doc->Save(dataDir);
    // ExEnd:RemoveFooters
    std::cout << "\nAll footers from all sections deleted successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
