#pragma once

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/CompressionLevel.h>
#include <Aspose.Words.Cpp/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Settings/MsWordVersion.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options {

class WorkingWithOoxmlSaveOptions : public DocsExamplesBase
{
public:
    void EncryptDocxWithPassword()
    {
        //ExStart:EncryptDocxWithPassword
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_Password(u"password");

        doc->Save(ArtifactsDir + u"WorkingWithOoxmlSaveOptions.EncryptDocxWithPassword.docx", saveOptions);
        //ExEnd:EncryptDocxWithPassword
    }

    void OoxmlComplianceIso29500_2008_Strict()
    {
        //ExStart:OoxmlComplianceIso29500_2008_Strict
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2016);

        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_Compliance(OoxmlCompliance::Iso29500_2008_Strict);

        doc->Save(ArtifactsDir + u"WorkingWithOoxmlSaveOptions.OoxmlComplianceIso29500_2008_Strict.docx", saveOptions);
        //ExEnd:OoxmlComplianceIso29500_2008_Strict
    }

    void UpdateLastSavedTimeProperty()
    {
        //ExStart:UpdateLastSavedTimeProperty
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_UpdateLastSavedTimeProperty(true);

        doc->Save(ArtifactsDir + u"WorkingWithOoxmlSaveOptions.UpdateLastSavedTimeProperty.docx", saveOptions);
        //ExEnd:UpdateLastSavedTimeProperty
    }

    void KeepLegacyControlChars()
    {
        //ExStart:KeepLegacyControlChars
        auto doc = MakeObject<Document>(MyDir + u"Legacy control character.doc");

        auto saveOptions = MakeObject<OoxmlSaveOptions>(SaveFormat::FlatOpc);
        saveOptions->set_KeepLegacyControlChars(true);

        doc->Save(ArtifactsDir + u"WorkingWithOoxmlSaveOptions.KeepLegacyControlChars.docx", saveOptions);
        //ExEnd:KeepLegacyControlChars
    }

    void SetCompressionLevel()
    {
        //ExStart:SetCompressionLevel
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_CompressionLevel(CompressionLevel::SuperFast);

        doc->Save(ArtifactsDir + u"WorkingWithOoxmlSaveOptions.SetCompressionLevel.docx", saveOptions);
        //ExEnd:SetCompressionLevel
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
