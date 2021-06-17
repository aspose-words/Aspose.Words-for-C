#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/EmphasisMark.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Fonts/DefaultFontSubstitutionRule.h>
#include <Aspose.Words.Cpp/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Fonts/FontFallbackSettings.h>
#include <Aspose.Words.Cpp/Fonts/FontInfoSubstitutionRule.h>
#include <Aspose.Words.Cpp/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Fonts/FontSubstitutionSettings.h>
#include <Aspose.Words.Cpp/Fonts/PhysicalFontInfo.h>
#include <Aspose.Words.Cpp/Fonts/StreamFontSource.h>
#include <Aspose.Words.Cpp/Fonts/SystemFontSource.h>
#include <Aspose.Words.Cpp/Fonts/TableSubstitutionRule.h>
#include <Aspose.Words.Cpp/IWarningCallback.h>
#include <Aspose.Words.Cpp/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/TextDmlEffect.h>
#include <Aspose.Words.Cpp/Underline.h>
#include <Aspose.Words.Cpp/WarningInfo.h>
#include <Aspose.Words.Cpp/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/WarningType.h>
#include <drawing/color.h>
#include <system/array.h>
#include <system/collections/ilist.h>
#include <system/collections/list.h>
#include <system/convert.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/io/stream.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Loading;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithFonts : public DocsExamplesBase
{
public:
    class HandleDocumentWarnings : public IWarningCallback
    {
    public:
        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// Potential issue during document procssing. The callback can be set to listen for warnings generated
        /// during document load and/or document save.
        /// </summary>
        void Warning(SharedPtr<WarningInfo> info) override
        {
            // We are only interested in fonts being substituted.
            if (info->get_WarningType() == WarningType::FontSubstitution)
            {
                std::cout << (String(u"Font substitution: ") + info->get_Description()) << std::endl;
            }
        }
    };

    class ResourceSteamFontSource : public StreamFontSource
    {
    public:
        SharedPtr<System::IO::Stream> OpenFontDataStream() override
        {
            return System::Reflection::Assembly::GetExecutingAssembly()->GetManifestResourceStream(u"resourceName");
        }

    protected:
        virtual ~ResourceSteamFontSource()
        {
        }
    };

    class DocumentSubstitutionWarnings : public IWarningCallback
    {
    public:
        SharedPtr<WarningInfoCollection> FontWarnings;

        /// <summary>
        /// Our callback only needs to implement the "Warning" method.
        /// This method is called whenever there is a potential issue during document processing.
        /// The callback can be set to listen for warnings generated during document load and/or document save.
        /// </summary>
        void Warning(SharedPtr<WarningInfo> info) override
        {
            // We are only interested in fonts being substituted.
            if (info->get_WarningType() == WarningType::FontSubstitution)
            {
                FontWarnings->Warning(info);
            }
        }

        DocumentSubstitutionWarnings() : FontWarnings(MakeObject<WarningInfoCollection>())
        {
        }
    };

public:
    void FontFormatting()
    {
        //ExStart:WriteAndFont
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Font> font = builder->get_Font();
        font->set_Size(16);
        font->set_Bold(true);
        font->set_Color(System::Drawing::Color::get_Blue());
        font->set_Name(u"Arial");
        font->set_Underline(Underline::Dash);

        builder->Write(u"Sample text.");

        doc->Save(ArtifactsDir + u"WorkingWithFonts.FontFormatting.docx");
        //ExEnd:WriteAndFont
    }

    void GetFontLineSpacing()
    {
        //ExStart:GetFontLineSpacing
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Calibri");
        builder->Writeln(u"qText");

        SharedPtr<Font> font = builder->get_Document()->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font();
        std::cout << "lineSpacing = " << font->get_LineSpacing() << std::endl;
        //ExEnd:GetFontLineSpacing
    }

    void CheckDMLTextEffect()
    {
        //ExStart:CheckDMLTextEffect
        auto doc = MakeObject<Document>(MyDir + u"DrawingML text effects.docx");

        SharedPtr<RunCollection> runs = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs();
        SharedPtr<Font> runFont = runs->idx_get(0)->get_Font();

        // One run might have several Dml text effects applied.
        std::cout << System::Convert::ToString(runFont->HasDmlEffect(TextDmlEffect::Shadow)) << std::endl;
        std::cout << System::Convert::ToString(runFont->HasDmlEffect(TextDmlEffect::Effect3D)) << std::endl;
        std::cout << System::Convert::ToString(runFont->HasDmlEffect(TextDmlEffect::Reflection)) << std::endl;
        std::cout << System::Convert::ToString(runFont->HasDmlEffect(TextDmlEffect::Outline)) << std::endl;
        std::cout << System::Convert::ToString(runFont->HasDmlEffect(TextDmlEffect::Fill)) << std::endl;
        //ExEnd:CheckDMLTextEffect
    }

    void SetFontFormatting()
    {
        //ExStart:DocumentBuilderSetFontFormatting
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Font> font = builder->get_Font();
        font->set_Bold(true);
        font->set_Color(System::Drawing::Color::get_DarkBlue());
        font->set_Italic(true);
        font->set_Name(u"Arial");
        font->set_Size(24);
        font->set_Spacing(5);
        font->set_Underline(Underline::Double);

        builder->Writeln(u"I'm a very nice formatted string.");

        doc->Save(ArtifactsDir + u"WorkingWithFonts.SetFontFormatting.docx");
        //ExEnd:DocumentBuilderSetFontFormatting
    }

    void SetFontEmphasisMark()
    {
        //ExStart:SetFontEmphasisMark
        auto document = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(document);

        builder->get_Font()->set_EmphasisMark(EmphasisMark::UnderSolidCircle);

        builder->Write(u"Emphasis text");
        builder->Writeln();
        builder->get_Font()->ClearFormatting();
        builder->Write(u"Simple text");

        document->Save(ArtifactsDir + u"WorkingWithFonts.SetFontEmphasisMark.docx");
        //ExEnd:SetFontEmphasisMark
    }

    void SetFontsFolders()
    {
        //ExStart:SetFontsFolders
        FontSettings::get_DefaultInstance()->SetFontsSources(
            MakeArray<SharedPtr<FontSourceBase>>({MakeObject<SystemFontSource>(), MakeObject<FolderFontSource>(u"C:\\MyFonts\\", true)}));

        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        doc->Save(ArtifactsDir + u"WorkingWithFonts.SetFontsFolders.pdf");
        //ExEnd:SetFontsFolders
    }

    void EnableDisableFontSubstitution()
    {
        //ExStart:EnableDisableFontSubstitution
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
        fontSettings->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(false);

        doc->set_FontSettings(fontSettings);

        doc->Save(ArtifactsDir + u"WorkingWithFonts.EnableDisableFontSubstitution.pdf");
        //ExEnd:EnableDisableFontSubstitution
    }

    void SetFontFallbackSettings()
    {
        //ExStart:SetFontFallbackSettings
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_FallbackSettings()->Load(MyDir + u"Font fallback rules.xml");

        doc->set_FontSettings(fontSettings);

        doc->Save(ArtifactsDir + u"WorkingWithFonts.SetFontFallbackSettings.pdf");
        //ExEnd:SetFontFallbackSettings
    }

    void NotoFallbackSettings()
    {
        //ExStart:SetPredefinedFontFallbackSettings
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_FallbackSettings()->LoadNotoFallbackSettings();

        doc->set_FontSettings(fontSettings);

        doc->Save(ArtifactsDir + u"WorkingWithFonts.NotoFallbackSettings.pdf");
        //ExEnd:SetPredefinedFontFallbackSettings
    }

    void SetFontsFoldersDefaultInstance()
    {
        //ExStart:SetFontsFoldersDefaultInstance
        FontSettings::get_DefaultInstance()->SetFontsFolder(u"C:\\MyFonts\\", true);
        //ExEnd:SetFontsFoldersDefaultInstance

        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        doc->Save(ArtifactsDir + u"WorkingWithFonts.SetFontsFoldersDefaultInstance.pdf");
    }

    void SetFontsFoldersMultipleFolders()
    {
        //ExStart:SetFontsFoldersMultipleFolders
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto fontSettings = MakeObject<FontSettings>();
        // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        // fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        // FontSettings.SetFontSources instead.
        fontSettings->SetFontsFolders(MakeArray<String>({u"C:\\MyFonts\\", u"D:\\Misc\\Fonts\\"}), true);

        doc->set_FontSettings(fontSettings);

        doc->Save(ArtifactsDir + u"WorkingWithFonts.SetFontsFoldersMultipleFolders.pdf");
        //ExEnd:SetFontsFoldersMultipleFolders
    }

    void SetFontsFoldersSystemAndCustomFolder()
    {
        //ExStart:SetFontsFoldersSystemAndCustomFolder
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto fontSettings = MakeObject<FontSettings>();
        // Retrieve the array of environment-dependent font sources that are searched by default.
        // For example this will contain a "Windows\Fonts\" source on a Windows machines.
        // We add this array to a new List to make adding or removing font entries much easier.
        SharedPtr<System::Collections::Generic::List<SharedPtr<FontSourceBase>>> fontSources =
            MakeObject<System::Collections::Generic::List<SharedPtr<FontSourceBase>>>(fontSettings->GetFontsSources());

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        auto folderFontSource = MakeObject<FolderFontSource>(u"C:\\MyFonts\\", true);

        // Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources->Add(folderFontSource);

        ArrayPtr<SharedPtr<FontSourceBase>> updatedFontSources = fontSources->ToArray();
        fontSettings->SetFontsSources(updatedFontSources);

        doc->set_FontSettings(fontSettings);

        doc->Save(ArtifactsDir + u"WorkingWithFonts.SetFontsFoldersSystemAndCustomFolder.pdf");
        //ExEnd:SetFontsFoldersSystemAndCustomFolder
    }

    void SetFontsFoldersWithPriority()
    {
        //ExStart:SetFontsFoldersWithPriority
        FontSettings::get_DefaultInstance()->SetFontsSources(
            MakeArray<SharedPtr<FontSourceBase>>({MakeObject<SystemFontSource>(), MakeObject<FolderFontSource>(u"C:\\MyFonts\\", true, 1)}));
        //ExEnd:SetFontsFoldersWithPriority

        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        doc->Save(ArtifactsDir + u"WorkingWithFonts.SetFontsFoldersWithPriority.pdf");
    }

    void SetTrueTypeFontsFolder()
    {
        //ExStart:SetTrueTypeFontsFolder
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto fontSettings = MakeObject<FontSettings>();
        // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        // Fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        // FontSettings.SetFontSources instead
        fontSettings->SetFontsFolder(u"C:\\MyFonts\\", false);
        // Set font settings
        doc->set_FontSettings(fontSettings);

        doc->Save(ArtifactsDir + u"WorkingWithFonts.SetTrueTypeFontsFolder.pdf");
        //ExEnd:SetTrueTypeFontsFolder
    }

    void SpecifyDefaultFontWhenRendering()
    {
        //ExStart:SpecifyDefaultFontWhenRendering
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto fontSettings = MakeObject<FontSettings>();
        // If the default font defined here cannot be found during rendering then
        // the closest font on the machine is used instead.
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial Unicode MS");

        doc->set_FontSettings(fontSettings);

        doc->Save(ArtifactsDir + u"WorkingWithFonts.SpecifyDefaultFontWhenRendering.pdf");
        //ExEnd:SpecifyDefaultFontWhenRendering
    }

    void FontSettingsWithLoadOptions()
    {
        //ExStart:FontSettingsWithLoadOptions
        auto fontSettings = MakeObject<FontSettings>();

        SharedPtr<TableSubstitutionRule> substitutionRule = fontSettings->get_SubstitutionSettings()->get_TableSubstitution();
        // If "UnknownFont1" font family is not available then substitute it by "Comic Sans MS"
        substitutionRule->AddSubstitutes(u"UnknownFont1", MakeArray<String>({u"Comic Sans MS"}));

        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_FontSettings(fontSettings);

        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx", loadOptions);
        //ExEnd:FontSettingsWithLoadOptions
    }

    void SetFontsFolder()
    {
        //ExStart:SetFontsFolder
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->SetFontsFolder(MyDir + u"Fonts", false);

        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_FontSettings(fontSettings);

        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx", loadOptions);
        //ExEnd:SetFontsFolder
    }

    void FontSettingsWithLoadOption()
    {
        //ExStart:FontSettingsWithLoadOption
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_FontSettings(MakeObject<FontSettings>());

        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx", loadOptions);
        //ExEnd:FontSettingsWithLoadOption
    }

    void FontSettingsDefaultInstance()
    {
        //ExStart:FontSettingsFontSource
        //ExStart:FontSettingsDefaultInstance
        SharedPtr<FontSettings> fontSettings = FontSettings::get_DefaultInstance();
        //ExEnd:FontSettingsDefaultInstance
        fontSettings->SetFontsSources(
            MakeArray<SharedPtr<FontSourceBase>>({MakeObject<SystemFontSource>(), MakeObject<FolderFontSource>(u"C:\\MyFonts\\", true)}));
        //ExEnd:FontSettingsFontSource

        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_FontSettings(fontSettings);
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx", loadOptions);
    }

    void GetListOfAvailableFonts()
    {
        //ExStart:GetListOfAvailableFonts
        auto fontSettings = MakeObject<FontSettings>();
        SharedPtr<System::Collections::Generic::List<SharedPtr<FontSourceBase>>> fontSources =
            MakeObject<System::Collections::Generic::List<SharedPtr<FontSourceBase>>>(fontSettings->GetFontsSources());

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        auto folderFontSource = MakeObject<FolderFontSource>(MyDir, true);
        // Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources->Add(folderFontSource);

        ArrayPtr<SharedPtr<FontSourceBase>> updatedFontSources = fontSources->ToArray();

        for (const auto& fontInfo : System::IterateOver(updatedFontSources[0]->GetAvailableFonts()))
        {
            std::cout << (String(u"FontFamilyName : ") + fontInfo->get_FontFamilyName()) << std::endl;
            std::cout << (String(u"FullFontName  : ") + fontInfo->get_FullFontName()) << std::endl;
            std::cout << (String(u"Version  : ") + fontInfo->get_Version()) << std::endl;
            std::cout << (String(u"FilePath : ") + fontInfo->get_FilePath()) << std::endl;
        }
        //ExEnd:GetListOfAvailableFonts
    }

    void ReceiveNotificationsOfFonts()
    {
        //ExStart:ReceiveNotificationsOfFonts
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto fontSettings = MakeObject<FontSettings>();

        // We can choose the default font to use in the case of any missing fonts.
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
        // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
        // find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default
        // font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback.
        fontSettings->SetFontsFolder(String::Empty, false);

        // Create a new class implementing IWarningCallback which collect any warnings produced during document save.
        auto callback = MakeObject<WorkingWithFonts::HandleDocumentWarnings>();

        doc->set_WarningCallback(callback);
        doc->set_FontSettings(fontSettings);

        doc->Save(ArtifactsDir + u"WorkingWithFonts.ReceiveNotificationsOfFonts.pdf");
        //ExEnd:ReceiveNotificationsOfFonts
    }

    void ReceiveWarningNotification()
    {
        //ExStart:ReceiveWarningNotification
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occured during rendering
        // are stored until the document save and then sent to the appropriate WarningCallback.
        doc->UpdatePageLayout();

        auto callback = MakeObject<WorkingWithFonts::HandleDocumentWarnings>();
        doc->set_WarningCallback(callback);

        // Even though the document was rendered previously, any save warnings are notified to the user during document save.
        doc->Save(ArtifactsDir + u"WorkingWithFonts.ReceiveWarningNotification.pdf");
        //ExEnd:ReceiveWarningNotification
    }

    void ResourceSteamFontSourceExample()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        FontSettings::get_DefaultInstance()->SetFontsSources(
            MakeArray<SharedPtr<FontSourceBase>>({MakeObject<SystemFontSource>(), MakeObject<WorkingWithFonts::ResourceSteamFontSource>()}));

        doc->Save(ArtifactsDir + u"WorkingWithFonts.SetFontsFolders.pdf");
    }

    void GetSubstitutionWithoutSuffixes()
    {
        auto doc = MakeObject<Document>(MyDir + u"Get substitution without suffixes.docx");

        auto substitutionWarningHandler = MakeObject<WorkingWithFonts::DocumentSubstitutionWarnings>();
        doc->set_WarningCallback(substitutionWarningHandler);

        SharedPtr<System::Collections::Generic::List<SharedPtr<FontSourceBase>>> fontSources =
            MakeObject<System::Collections::Generic::List<SharedPtr<FontSourceBase>>>(FontSettings::get_DefaultInstance()->GetFontsSources());

        auto folderFontSource = MakeObject<FolderFontSource>(FontsDir, true);
        fontSources->Add(folderFontSource);

        ArrayPtr<SharedPtr<FontSourceBase>> updatedFontSources = fontSources->ToArray();
        FontSettings::get_DefaultInstance()->SetFontsSources(updatedFontSources);

        doc->Save(ArtifactsDir + u"WorkingWithFonts.GetSubstitutionWithoutSuffixes.pdf");

        ASSERT_EQ(u"Font 'DINOT-Regular' has not been found. Using 'DINOT' font instead. Reason: font name substitution.",
                  substitutionWarningHandler->FontWarnings->idx_get(0)->get_Description());
    }
};

}} // namespace DocsExamples::Programming_with_Documents
