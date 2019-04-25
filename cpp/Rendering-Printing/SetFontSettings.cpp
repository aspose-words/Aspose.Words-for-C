#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Fonts/DefaultFontSubstitutionRule.h>
#include <Model/Fonts/FontInfoSubstitutionRule.h>
#include <Model/Fonts/FontFallbackSettings.h>
#include <Model/Fonts/FontSettings.h>
#include <Model/Fonts/FontSubstitutionSettings.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;

namespace
{
    void EnableDisableFontSubstitution(System::String const &dataDir)
    {
        // ExStart:EnableDisableFontSubstitution
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Rendering.doc");

        System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>();
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
        fontSettings->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(false);

        // Set font settings
        doc->set_FontSettings(fontSettings);
        System::String outputPath = dataDir + GetOutputFilePath(u"SetFontSettings.EnableDisableFontSubstitution.pdf");
        doc->Save(outputPath);
        // ExEnd:EnableDisableFontSubstitution
        std::cout << "Document is rendered to PDF with disabled font substitution." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetFontFallbackSettings(System::String const &dataDir)
    {
        // ExStart:SetFontFallbackSettings
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Rendering.doc");

        System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>();
        fontSettings->get_FallbackSettings()->Load(dataDir + u"Fallback.xml");

        // Set font settings
        doc->set_FontSettings(fontSettings);
        System::String outputPath = dataDir + GetOutputFilePath(u"SetFontSettings.SetFontFallbackSettings.pdf");
        doc->Save(outputPath);
        // ExEnd:SetFontFallbackSettings
        std::cout << "Document is rendered to PDF with font fallback." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetPredefinedFontFallbackSettings(System::String const &dataDir)
    {
        // ExStart:SetPredefinedFontFallbackSettings
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Rendering.doc");

        System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>();
        fontSettings->get_FallbackSettings()->LoadNotoFallbackSettings();

        // Set font settings
        doc->set_FontSettings(fontSettings);
        System::String outputPath = dataDir + GetOutputFilePath(u"SetFontSettings.SetPredefinedFontFallbackSettings.pdf");
        doc->Save(outputPath);
        // ExEnd:SetPredefinedFontFallbackSettings
        std::cout << "Document is rendered to PDF with font fallback." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void SetFontSettings()
{
    std::cout << "SetFontSettings example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_RenderingAndPrinting();
    EnableDisableFontSubstitution(dataDir);
    SetFontFallbackSettings(dataDir);
    // TODO (std_string) : FontFallbackSettings::LoadNotoFallbackSettings() don't work properly
    SetPredefinedFontFallbackSettings(dataDir);
    std::cout << "SetFontSettings example finished." << std::endl << std::endl;
}