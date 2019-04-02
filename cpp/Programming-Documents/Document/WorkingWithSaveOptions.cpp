#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Saving/OoxmlSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void UpdateLastSavedTimeProperty(System::String const &dataDir)
    {
        // ExStart:UpdateLastSavedTimeProperty
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");

        System::SharedPtr<OoxmlSaveOptions> options = System::MakeObject<OoxmlSaveOptions>();
        options->set_UpdateLastSavedTimeProperty(true);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithSaveOptions.UpdateLastSavedTimeProperty.docx");

        // Save the document to disk.
        doc->Save(outputPath, options);
        // ExEnd:UpdateLastSavedTimeProperty
        std::cout << "Updated Last Saved Time Property successfully." << std::endl;
    }

}

void WorkingWithSaveOptions()
{
    std::cout << "WorkingWithSaveOptions example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    UpdateLastSavedTimeProperty(dataDir);
    std::cout << "WorkingWithSaveOptions example finished." << std::endl << std::endl;
}