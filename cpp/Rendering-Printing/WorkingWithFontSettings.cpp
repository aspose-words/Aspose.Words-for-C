#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Fonts/DefaultFontSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontInfoSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontFallbackSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSubstitutionSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/TableSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/SystemFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontConfigSubstitutionRule.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;

namespace
{
    void FontSettingsWithLoadOptions(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:FontSettingsWithLoadOptions
        System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>();
        System::SharedPtr<TableSubstitutionRule> substitutionRule = fontSettings->get_SubstitutionSettings()->get_TableSubstitution();
        
        // If "UnknownFont1" font family is not available then substitute it by "Comic Sans MS".
        substitutionRule->AddSubstitutes(u"UnknownFont1", System::MakeArray<System::String>({ u"Comic Sans MS" }));
        
        System::SharedPtr<LoadOptions> lo = System::MakeObject<LoadOptions>();
        lo->set_FontSettings(fontSettings);
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"MyDocument.docx", lo);
        // ExEnd:FontSettingsWithLoadOptions
    }
    void EnableDisableFontSubstitution(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:EnableDisableFontSubstitution
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

        System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>();
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
        fontSettings->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(false);

        // Set font settings
        doc->set_FontSettings(fontSettings);
        System::String outputPath = outputDataDir + u"SetFontSettings.EnableDisableFontSubstitution.pdf";
        doc->Save(outputPath);
        // ExEnd:EnableDisableFontSubstitution
        std::cout << "Document is rendered to PDF with disabled font substitution." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetFontFallbackSettings(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetFontFallbackSettings
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

        System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>();
        fontSettings->get_FallbackSettings()->Load(inputDataDir + u"Fallback.xml");

        // Set font settings
        doc->set_FontSettings(fontSettings);
        System::String outputPath = outputDataDir + u"SetFontSettings.SetFontFallbackSettings.pdf";
        doc->Save(outputPath);
        // ExEnd:SetFontFallbackSettings
        std::cout << "Document is rendered to PDF with font fallback." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetPredefinedFontFallbackSettings(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetPredefinedFontFallbackSettings
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

        System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>();
        fontSettings->get_FallbackSettings()->LoadNotoFallbackSettings();

        // Set font settings
        doc->set_FontSettings(fontSettings);
        System::String outputPath = outputDataDir + u"SetFontSettings.SetPredefinedFontFallbackSettings.pdf";
        doc->Save(outputPath);
        // ExEnd:SetPredefinedFontFallbackSettings
        std::cout << "Document is rendered to PDF with font fallback." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void FontSettingsDefaultInstance(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:FontSettingsFontSource
        // ExStart:FontSettingsDefaultInstance
        System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>()->get_DefaultInstance();
        // ExEnd:FontSettingsDefaultInstance   
        fontSettings->SetFontsSources(System::MakeArray<System::SharedPtr<FontSourceBase>>(
            {
                System::MakeObject<SystemFontSource>(),
                System::MakeObject<FolderFontSource>(u"/home/user/MyFonts", true)
            }));
        // ExEnd:FontSettingsFontSource

        // init font settings
        System::SharedPtr<LoadOptions> lo = System::MakeObject<LoadOptions>();
        lo->set_FontSettings(fontSettings);
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"MyDocument.docx", lo);
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
        System::SharedPtr<FontFallbackSettings> settings = fontSettings->get_FallbackSettings();
        
    }

}

void WorkingWithFontSettings()
{
    std::cout << "SetFontSettings example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();
    EnableDisableFontSubstitution(inputDataDir, outputDataDir);
    SetFontFallbackSettings(inputDataDir, outputDataDir);
    SetPredefinedFontFallbackSettings(inputDataDir, outputDataDir);
    std::cout << "SetFontSettings example finished." << std::endl << std::endl;
}