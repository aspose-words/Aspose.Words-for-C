#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Saving/OoxmlCompliance.h>
#include <Model/Saving/OoxmlSaveOptions.h>
#include <Model/Settings/CompatibilityOptions.h>
#include <Model/Settings/MsWordVersion.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace
{
    void SetOOXMLCompliance(System::String const &dataDir)
    {
        //ExStart:SetOOXMLCompliance
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");

        // Set Word2016 version for document
        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2016);

        //Set the Strict compliance level. 
        System::SharedPtr<OoxmlSaveOptions> ooxmlSaveOptions = System::MakeObject<OoxmlSaveOptions>();
        ooxmlSaveOptions->set_Compliance(OoxmlCompliance::Iso29500_2008_Strict);
        ooxmlSaveOptions->set_SaveFormat(SaveFormat::Docx);
        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithOoxml.SetOOXMLCompliance.docx");
        doc->Save(outputPath, ooxmlSaveOptions);
        //ExEnd:SetOOXMLCompliance
        std::cout << "Document is saved with ISO/IEC 29500:2008 Strict compliance level." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void UpdateLastSavedTimeProperty(System::String const &dataDir)
    {
        // ExStart:UpdateLastSavedTimeProperty
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");

        System::SharedPtr<OoxmlSaveOptions> ooxmlSaveOptions = System::MakeObject<OoxmlSaveOptions>();
        ooxmlSaveOptions->set_UpdateLastSavedTimeProperty(true);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithOoxml.UpdateLastSavedTimeProperty.docx");
        // Save the document to disk.
        doc->Save(outputPath, ooxmlSaveOptions);
        // ExEnd:UpdateLastSavedTimeProperty
        std::cout << "Updated Last Saved Time Property successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void KeepLegacyControlChars(System::String const &dataDir)
    {
        // ExStart:KeepLegacyControlChars
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");

        System::SharedPtr<OoxmlSaveOptions> so = System::MakeObject<OoxmlSaveOptions>(SaveFormat::FlatOpc);
        so->set_KeepLegacyControlChars(true);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithOoxml.KeepLegacyControlChars.docx");
        // Save the document to disk.
        doc->Save(outputPath, so);

        // ExEnd:KeepLegacyControlChars
        std::cout << "Updated Last Saved With Keeping Legacy Control Chars Successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkingWithOoxml()
{
    std::cout << "WorkingWithOoxml example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();
    SetOOXMLCompliance(dataDir);
    UpdateLastSavedTimeProperty(dataDir);
    KeepLegacyControlChars(dataDir);
    std::cout << "WorkingWithOoxml example finished." << std::endl << std::endl;
}