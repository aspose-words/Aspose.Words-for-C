#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Fonts/DefaultFontSubstitutionRule.h>
#include <Aspose.Words.Cpp/Fonts/FileFontSource.h>
#include <Aspose.Words.Cpp/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Fonts/FontConfigSubstitutionRule.h>
#include <Aspose.Words.Cpp/Fonts/FontFallbackSettings.h>
#include <Aspose.Words.Cpp/Fonts/FontInfoSubstitutionRule.h>
#include <Aspose.Words.Cpp/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Fonts/FontSourceType.h>
#include <Aspose.Words.Cpp/Fonts/FontSubstitutionSettings.h>
#include <Aspose.Words.Cpp/Fonts/MemoryFontSource.h>
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
#include <Aspose.Words.Cpp/WarningInfo.h>
#include <Aspose.Words.Cpp/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/WarningSource.h>
#include <Aspose.Words.Cpp/WarningType.h>
#include <net/http_status_code.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ilist.h>
#include <system/collections/list.h>
#include <system/convert.h>
#include <system/details/dispose_guard.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/environment.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/io/file.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/operating_system.h>
#include <system/platform_id.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <xml/xml_attribute.h>
#include <xml/xml_attribute_collection.h>
#include <xml/xml_document.h>
#include <xml/xml_name_table.h>
#include <xml/xml_namespace_manager.h>
#include <xml/xml_node.h>
#include <xml/xml_node_list.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Loading;

namespace ApiExamples {

class ExFontSettings : public ApiExampleBase
{
public:
    void DefaultFontInstance()
    {
        //ExStart
        //ExFor:Fonts.FontSettings.DefaultInstance
        //ExSummary:Shows how to configure the default font settings instance.
        // Configure the default font settings instance to use the "Courier New" font
        // as a backup substitute when we attempt to use an unknown font.
        FontSettings::get_DefaultInstance()->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Courier New");

        ASSERT_TRUE(FontSettings::get_DefaultInstance()->get_SubstitutionSettings()->get_DefaultFontSubstitution()->get_Enabled());

        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Non-existent font");
        builder->Write(u"Hello world!");

        // This document does not have a FontSettings configuration. When we render the document,
        // the default FontSettings instance will resolve the missing font.
        // Aspose.Words will use "Courier New" to render text that uses the unknown font.
        ASSERT_TRUE(doc->get_FontSettings() == nullptr);

        doc->Save(ArtifactsDir + u"FontSettings.DefaultFontInstance.pdf");
        //ExEnd
    }

    void DefaultFontName()
    {
        //ExStart
        //ExFor:DefaultFontSubstitutionRule.DefaultFontName
        //ExSummary:Shows how to specify a default font.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(u"Hello world!");
        builder->get_Font()->set_Name(u"Arvo");
        builder->Writeln(u"The quick brown fox jumps over the lazy dog.");

        ArrayPtr<SharedPtr<FontSourceBase>> fontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        // The font sources that the document uses contain the font "Arial", but not "Arvo".
        ASSERT_EQ(1, fontSources->get_Length());
        ASSERT_TRUE(fontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arial"; }));
        ASSERT_FALSE(fontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arvo"; }));

        // Set the "DefaultFontName" property to "Courier New" to,
        // while rendering the document, apply that font in all cases when another font is not available.
        FontSettings::get_DefaultInstance()->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Courier New");

        ASSERT_TRUE(fontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Courier New"; }));

        // Aspose.Words will now use the default font in place of any missing fonts during any rendering calls.
        doc->Save(ArtifactsDir + u"FontSettings.DefaultFontName.pdf");
        //ExEnd
    }

    void UpdatePageLayoutWarnings()
    {
        // Store the font sources currently used so we can restore them later
        ArrayPtr<SharedPtr<FontSourceBase>> originalFontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        // Load the document to render
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
        auto callback = MakeObject<ExFontSettings::HandleDocumentWarnings>();
        doc->set_WarningCallback(callback);

        // We can choose the default font to use in the case of any missing fonts
        FontSettings::get_DefaultInstance()->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");

        // For testing we will set Aspose.Words to look for fonts only in a folder which does not exist. Since Aspose.Words won't
        // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default
        // font specified under FontSettings.DefaultFontName. We can pick up on this substitution using our callback
        FontSettings::get_DefaultInstance()->SetFontsFolder(String::Empty, false);

        // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occurred during rendering
        // are stored until the document save and then sent to the appropriate WarningCallback
        doc->UpdatePageLayout();

        // Even though the document was rendered previously, any save warnings are notified to the user during document save
        doc->Save(ArtifactsDir + u"FontSettings.UpdatePageLayoutWarnings.pdf");

        ASSERT_GT(callback->FontWarnings->get_Count(), 0);
        ASSERT_TRUE(callback->FontWarnings->idx_get(0)->get_WarningType() == WarningType::FontSubstitution);
        ASSERT_TRUE(callback->FontWarnings->idx_get(0)->get_Description().Contains(u"has not been found"));

        // Restore default fonts
        FontSettings::get_DefaultInstance()->SetFontsSources(originalFontSources);
    }

    class HandleDocumentWarnings : public IWarningCallback
    {
    public:
        SharedPtr<WarningInfoCollection> FontWarnings;

        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        void Warning(SharedPtr<WarningInfo> info) override
        {
            // We are only interested in fonts being substituted
            if (info->get_WarningType() == WarningType::FontSubstitution)
            {
                std::cout << (String(u"Font substitution: ") + info->get_Description()) << std::endl;
                FontWarnings->Warning(info);
            }
        }

        HandleDocumentWarnings() : FontWarnings(MakeObject<WarningInfoCollection>())
        {
        }
    };

    //ExStart
    //ExFor:IWarningCallback
    //ExFor:DocumentBase.WarningCallback
    //ExFor:Fonts.FontSettings.DefaultInstance
    //ExSummary:Shows how to use the IWarningCallback interface to monitor font substitution warnings.
    void SubstitutionWarning()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Times New Roman");
        builder->Writeln(u"Hello world!");

        auto callback = MakeObject<ExFontSettings::FontSubstitutionWarningCollector>();
        doc->set_WarningCallback(callback);

        // Store the current collection of font sources, which will be the default font source for every document
        // for which we do not specify a different font source.
        ArrayPtr<SharedPtr<FontSourceBase>> originalFontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        // For testing purposes, we will set Aspose.Words to look for fonts only in a folder that does not exist.
        FontSettings::get_DefaultInstance()->SetFontsFolder(String::Empty, false);

        // When rendering the document, there will be no place to find the "Times New Roman" font.
        // This will cause a font substitution warning, which our callback will detect.
        doc->Save(ArtifactsDir + u"FontSettings.SubstitutionWarning.pdf");

        FontSettings::get_DefaultInstance()->SetFontsSources(originalFontSources);

        ASSERT_EQ(1, callback->FontSubstitutionWarnings->get_Count());
        //ExSkip
        ASSERT_TRUE(callback->FontSubstitutionWarnings->idx_get(0)->get_WarningType() == WarningType::FontSubstitution);
        ASSERT_TRUE(System::ObjectExt::Equals(callback->FontSubstitutionWarnings->idx_get(0)->get_Description(),
                                              u"Font 'Times New Roman' has not been found. Using 'Fanwood' font instead. Reason: first available font."));
    }

    class FontSubstitutionWarningCollector : public IWarningCallback
    {
    public:
        SharedPtr<WarningInfoCollection> FontSubstitutionWarnings;

        /// <summary>
        /// Called every time a warning occurs during loading/saving.
        /// </summary>
        void Warning(SharedPtr<WarningInfo> info) override
        {
            if (info->get_WarningType() == WarningType::FontSubstitution)
            {
                FontSubstitutionWarnings->Warning(info);
            }
        }

        FontSubstitutionWarningCollector() : FontSubstitutionWarnings(MakeObject<WarningInfoCollection>())
        {
        }
    };
    //ExEnd

    //ExStart
    //ExFor:FontSourceBase.WarningCallback
    //ExSummary:Shows how to call warning callback when the font sources working with.
    void FontSourceWarning()
    {
        auto settings = MakeObject<FontSettings>();
        settings->SetFontsFolder(u"bad folder?", false);

        SharedPtr<FontSourceBase> source = settings->GetFontsSources()->idx_get(0);
        auto callback = MakeObject<ExFontSettings::FontSourceWarningCollector>();
        source->set_WarningCallback(callback);

        // Get the list of fonts to call warning callback.
        SharedPtr<System::Collections::Generic::IList<SharedPtr<PhysicalFontInfo>>> fontInfos = source->GetAvailableFonts();

        ASSERT_TRUE(callback->FontSubstitutionWarnings->idx_get(0)->get_Description().Contains(u"Error loading font from the folder \"bad folder?\""));
    }

    class FontSourceWarningCollector : public IWarningCallback
    {
    public:
        SharedPtr<WarningInfoCollection> FontSubstitutionWarnings;

        /// <summary>
        /// Called every time a warning occurs during processing of font source.
        /// </summary>
        void Warning(SharedPtr<WarningInfo> info) override
        {
            FontSubstitutionWarnings->Warning(info);
        }

        FontSourceWarningCollector() : FontSubstitutionWarnings(MakeObject<WarningInfoCollection>())
        {
        }
    };
    //ExEnd

    //ExStart
    //ExFor:Fonts.FontInfoSubstitutionRule
    //ExFor:Fonts.FontSubstitutionSettings.FontInfoSubstitution
    //ExFor:IWarningCallback
    //ExFor:IWarningCallback.Warning(WarningInfo)
    //ExFor:WarningInfo
    //ExFor:WarningInfo.Description
    //ExFor:WarningInfo.WarningType
    //ExFor:WarningInfoCollection
    //ExFor:WarningInfoCollection.Warning(WarningInfo)
    //ExFor:WarningInfoCollection.GetEnumerator
    //ExFor:WarningInfoCollection.Clear
    //ExFor:WarningType
    //ExFor:DocumentBase.WarningCallback
    //ExSummary:Shows how to set the property for finding the closest match for a missing font from the available font sources.
    void EnableFontSubstitution()
    {
        // Open a document that contains text formatted with a font that does not exist in any of our font sources.
        auto doc = MakeObject<Document>(MyDir + u"Missing font.docx");

        // Assign a callback for handling font substitution warnings.
        auto substitutionWarningHandler = MakeObject<ExFontSettings::HandleDocumentSubstitutionWarnings>();
        doc->set_WarningCallback(substitutionWarningHandler);

        // Set a default font name and enable font substitution.
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
        ;
        fontSettings->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(true);

        // We will get a font substitution warning if we save a document with a missing font.
        doc->set_FontSettings(fontSettings);
        doc->Save(ArtifactsDir + u"FontSettings.EnableFontSubstitution.pdf");

        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<WarningInfo>>> warnings = substitutionWarningHandler->FontWarnings->GetEnumerator();
            while (warnings->MoveNext())
            {
                std::cout << warnings->get_Current()->get_Description() << std::endl;
            }
        }

        // We can also verify warnings in the collection and clear them.
        ASSERT_EQ(WarningSource::Layout, substitutionWarningHandler->FontWarnings->idx_get(0)->get_Source());
        ASSERT_EQ(u"Font '28 Days Later' has not been found. Using 'Calibri' font instead. Reason: alternative name from document.",
                  substitutionWarningHandler->FontWarnings->idx_get(0)->get_Description());

        substitutionWarningHandler->FontWarnings->Clear();

        ASSERT_EQ(0, substitutionWarningHandler->FontWarnings->get_Count());
    }

    class HandleDocumentSubstitutionWarnings : public IWarningCallback
    {
    public:
        SharedPtr<WarningInfoCollection> FontWarnings;

        /// <summary>
        /// Called every time a warning occurs during loading/saving.
        /// </summary>
        void Warning(SharedPtr<WarningInfo> info) override
        {
            if (info->get_WarningType() == WarningType::FontSubstitution)
            {
                FontWarnings->Warning(info);
            }
        }

        HandleDocumentSubstitutionWarnings() : FontWarnings(MakeObject<WarningInfoCollection>())
        {
        }
    };
    //ExEnd

    void SubstitutionWarningsClosestMatch()
    {
        auto doc = MakeObject<Document>(MyDir + u"Bullet points with alternative font.docx");

        auto callback = MakeObject<ExFontSettings::HandleDocumentSubstitutionWarnings>();
        doc->set_WarningCallback(callback);

        doc->Save(ArtifactsDir + u"FontSettings.SubstitutionWarningsClosestMatch.pdf");

        ASSERT_TRUE(System::ObjectExt::Equals(callback->FontWarnings->idx_get(0)->get_Description(),
                                              u"Font \'SymbolPS\' has not been found. Using \'Wingdings\' font instead. Reason: font info substitution."));
    }

    void DisableFontSubstitution()
    {
        auto doc = MakeObject<Document>(MyDir + u"Missing font.docx");

        auto callback = MakeObject<ExFontSettings::HandleDocumentSubstitutionWarnings>();
        doc->set_WarningCallback(callback);

        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
        fontSettings->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(false);

        doc->set_FontSettings(fontSettings);
        doc->Save(ArtifactsDir + u"FontSettings.DisableFontSubstitution.pdf");

        auto reg = MakeObject<System::Text::RegularExpressions::Regex>(
            u"Font '28 Days Later' has not been found. Using (.*) font instead. Reason: default font setting.");

        for (const auto& fontWarning : callback->FontWarnings)
        {
            SharedPtr<System::Text::RegularExpressions::Match> match = reg->Match(fontWarning->get_Description());
            if (match->get_Success())
            {
                SUCCEED();
                return;
                break;
            }
        }
    }

    void SubstitutionWarnings()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto callback = MakeObject<ExFontSettings::HandleDocumentSubstitutionWarnings>();
        doc->set_WarningCallback(callback);

        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
        fontSettings->SetFontsFolder(FontsDir, false);
        fontSettings->get_SubstitutionSettings()->get_TableSubstitution()->AddSubstitutes(u"Arial", MakeArray<String>({u"Arvo", u"Slab"}));

        doc->set_FontSettings(fontSettings);
        doc->Save(ArtifactsDir + u"FontSettings.SubstitutionWarnings.pdf");

        ASSERT_EQ(u"Font \'Arial\' has not been found. Using \'Arvo\' font instead. Reason: table substitution.",
                  callback->FontWarnings->idx_get(0)->get_Description());
        ASSERT_EQ(u"Font \'Times New Roman\' has not been found. Using \'M+ 2m\' font instead. Reason: font info substitution.",
                  callback->FontWarnings->idx_get(1)->get_Description());
    }

    void GetSubstitutionWithoutSuffixes()
    {
        auto doc = MakeObject<Document>(MyDir + u"Get substitution without suffixes.docx");

        ArrayPtr<SharedPtr<FontSourceBase>> originalFontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        auto substitutionWarningHandler = MakeObject<ExFontSettings::HandleDocumentSubstitutionWarnings>();
        doc->set_WarningCallback(substitutionWarningHandler);

        SharedPtr<System::Collections::Generic::List<SharedPtr<FontSourceBase>>> fontSources =
            MakeObject<System::Collections::Generic::List<SharedPtr<FontSourceBase>>>(FontSettings::get_DefaultInstance()->GetFontsSources());
        auto folderFontSource = MakeObject<FolderFontSource>(FontsDir, true);
        fontSources->Add(folderFontSource);

        ArrayPtr<SharedPtr<FontSourceBase>> updatedFontSources = fontSources->ToArray();
        FontSettings::get_DefaultInstance()->SetFontsSources(updatedFontSources);

        doc->Save(ArtifactsDir + u"Font.GetSubstitutionWithoutSuffixes.pdf");

        ASSERT_EQ(u"Font 'DINOT-Regular' has not been found. Using 'DINOT' font instead. Reason: font name substitution.",
                  substitutionWarningHandler->FontWarnings->idx_get(0)->get_Description());

        FontSettings::get_DefaultInstance()->SetFontsSources(originalFontSources);
    }

    void FontSourceFile()
    {
        //ExStart
        //ExFor:Fonts.FileFontSource
        //ExFor:Fonts.FileFontSource.#ctor(String)
        //ExFor:Fonts.FileFontSource.#ctor(String, Int32)
        //ExFor:Fonts.FileFontSource.FilePath
        //ExFor:Fonts.FileFontSource.Type
        //ExFor:Fonts.FontSourceBase
        //ExFor:Fonts.FontSourceBase.Priority
        //ExFor:Fonts.FontSourceBase.Type
        //ExFor:Fonts.FontSourceType
        //ExSummary:Shows how to use a font file in the local file system as a font source.
        auto fileFontSource = MakeObject<FileFontSource>(MyDir + u"Alte DIN 1451 Mittelschrift.ttf", 0);

        auto doc = MakeObject<Document>();
        doc->set_FontSettings(MakeObject<FontSettings>());
        doc->get_FontSettings()->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({fileFontSource}));

        ASSERT_EQ(MyDir + u"Alte DIN 1451 Mittelschrift.ttf", fileFontSource->get_FilePath());
        ASSERT_EQ(FontSourceType::FontFile, fileFontSource->get_Type());
        ASSERT_EQ(0, fileFontSource->get_Priority());
        //ExEnd
    }

    void FontSourceFolder()
    {
        //ExStart
        //ExFor:Fonts.FolderFontSource
        //ExFor:Fonts.FolderFontSource.#ctor(String, Boolean)
        //ExFor:Fonts.FolderFontSource.#ctor(String, Boolean, Int32)
        //ExFor:Fonts.FolderFontSource.FolderPath
        //ExFor:Fonts.FolderFontSource.ScanSubfolders
        //ExFor:Fonts.FolderFontSource.Type
        //ExSummary:Shows how to use a local system folder which contains fonts as a font source.

        // Create a font source from a folder that contains font files.
        auto folderFontSource = MakeObject<FolderFontSource>(FontsDir, false, 1);

        auto doc = MakeObject<Document>();
        doc->set_FontSettings(MakeObject<FontSettings>());
        doc->get_FontSettings()->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({folderFontSource}));

        ASSERT_EQ(FontsDir, folderFontSource->get_FolderPath());
        ASPOSE_ASSERT_EQ(false, folderFontSource->get_ScanSubfolders());
        ASSERT_EQ(FontSourceType::FontsFolder, folderFontSource->get_Type());
        ASSERT_EQ(1, folderFontSource->get_Priority());
        //ExEnd
    }

    void SetFontsFolder(bool recursive)
    {
        //ExStart
        //ExFor:FontSettings
        //ExFor:FontSettings.SetFontsFolder(String, Boolean)
        //ExSummary:Shows how to set a font source directory.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Arvo");
        builder->Writeln(u"Hello world!");
        builder->get_Font()->set_Name(u"Amethysta");
        builder->Writeln(u"The quick brown fox jumps over the lazy dog.");

        // Our font sources do not contain the font that we have used for text in this document.
        // If we use these font settings while rendering this document,
        // Aspose.Words will apply a fallback font to text which has a font that Aspose.Words cannot locate.
        ArrayPtr<SharedPtr<FontSourceBase>> originalFontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        ASSERT_EQ(1, originalFontSources->get_Length());
        ASSERT_TRUE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arial"; }));

        // The default font sources are missing the two fonts that we are using in this document.
        ASSERT_FALSE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arvo"; }));
        ASSERT_FALSE(
            originalFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Amethysta"; }));

        // Use the "SetFontsFolder" method to set a directory which will act as a new font source.
        // Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directory
        // that we are passing in the first argument, but not include any fonts in any of that directory's subfolders.
        // Pass "true" as the "recursive" argument to include all font files in the directory that we are passing
        // in the first argument, as well as all the fonts in its subdirectories.
        FontSettings::get_DefaultInstance()->SetFontsFolder(FontsDir, recursive);

        ArrayPtr<SharedPtr<FontSourceBase>> newFontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        ASSERT_EQ(1, newFontSources->get_Length());
        ASSERT_FALSE(newFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arial"; }));
        ASSERT_TRUE(newFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arvo"; }));

        // The "Amethysta" font is in a subfolder of the font directory.
        if (recursive)
        {
            ASSERT_EQ(25, newFontSources[0]->GetAvailableFonts()->get_Count());
            ASSERT_TRUE(newFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Amethysta"; }));
        }
        else
        {
            ASSERT_EQ(18, newFontSources[0]->GetAvailableFonts()->get_Count());
            ASSERT_FALSE(newFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Amethysta"; }));
        }

        doc->Save(ArtifactsDir + u"FontSettings.SetFontsFolder.pdf");

        // Restore the original font sources.
        FontSettings::get_DefaultInstance()->SetFontsSources(originalFontSources);
        //ExEnd
    }

    void SetFontsFolders(bool recursive)
    {
        //ExStart
        //ExFor:FontSettings
        //ExFor:FontSettings.SetFontsFolders(String[], Boolean)
        //ExSummary:Shows how to set multiple font source directories.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Amethysta");
        builder->Writeln(u"The quick brown fox jumps over the lazy dog.");
        builder->get_Font()->set_Name(u"Junction Light");
        builder->Writeln(u"The quick brown fox jumps over the lazy dog.");

        // Our font sources do not contain the font that we have used for text in this document.
        // If we use these font settings while rendering this document,
        // Aspose.Words will apply a fallback font to text which has a font that Aspose.Words cannot locate.
        ArrayPtr<SharedPtr<FontSourceBase>> originalFontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        ASSERT_EQ(1, originalFontSources->get_Length());
        ASSERT_TRUE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arial"; }));

        // The default font sources are missing the two fonts that we are using in this document.
        ASSERT_FALSE(
            originalFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Amethysta"; }));
        ASSERT_FALSE(
            originalFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Junction Light"; }));

        // Use the "SetFontsFolders" method to create a font source from each font directory that we pass as the first argument.
        // Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directories
        // that we are passing in the first argument, but not include any fonts from any of the directories' subfolders.
        // Pass "true" as the "recursive" argument to include all font files in the directories that we are passing
        // in the first argument, as well as all the fonts in their subdirectories.
        FontSettings::get_DefaultInstance()->SetFontsFolders(MakeArray<String>({FontsDir + u"/Amethysta", FontsDir + u"/Junction"}), recursive);

        ArrayPtr<SharedPtr<FontSourceBase>> newFontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        ASSERT_EQ(2, newFontSources->get_Length());
        ASSERT_FALSE(newFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arial"; }));
        ASSERT_EQ(1, newFontSources[0]->GetAvailableFonts()->get_Count());
        ASSERT_TRUE(newFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Amethysta"; }));

        // The "Junction" folder itself contains no font files, but has subfolders that do.
        if (recursive)
        {
            ASSERT_EQ(6, newFontSources[1]->GetAvailableFonts()->get_Count());
            ASSERT_TRUE(
                newFontSources[1]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Junction Light"; }));
        }
        else
        {
            ASSERT_EQ(0, newFontSources[1]->GetAvailableFonts()->get_Count());
        }

        doc->Save(ArtifactsDir + u"FontSettings.SetFontsFolders.pdf");

        // Restore the original font sources.
        FontSettings::get_DefaultInstance()->SetFontsSources(originalFontSources);
        //ExEnd
    }

    void AddFontSource()
    {
        //ExStart
        //ExFor:FontSettings
        //ExFor:FontSettings.GetFontsSources()
        //ExFor:FontSettings.SetFontsSources(FontSourceBase[])
        //ExSummary:Shows how to add a font source to our existing font sources.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(u"Hello world!");
        builder->get_Font()->set_Name(u"Amethysta");
        builder->Writeln(u"The quick brown fox jumps over the lazy dog.");
        builder->get_Font()->set_Name(u"Junction Light");
        builder->Writeln(u"The quick brown fox jumps over the lazy dog.");

        ArrayPtr<SharedPtr<FontSourceBase>> originalFontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        ASSERT_EQ(1, originalFontSources->get_Length());

        ASSERT_TRUE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arial"; }));

        // The default font source is missing two of the fonts that we are using in our document.
        // When we save this document, Aspose.Words will apply fallback fonts to all text formatted with inaccessible fonts.
        ASSERT_FALSE(
            originalFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Amethysta"; }));
        ASSERT_FALSE(
            originalFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Junction Light"; }));

        // Create a font source from a folder that contains fonts.
        auto folderFontSource = MakeObject<FolderFontSource>(FontsDir, true);

        // Apply a new array of font sources that contains the original font sources, as well as our custom fonts.
        ArrayPtr<SharedPtr<FontSourceBase>> updatedFontSources = MakeArray<SharedPtr<FontSourceBase>>({originalFontSources[0], folderFontSource});
        FontSettings::get_DefaultInstance()->SetFontsSources(updatedFontSources);

        // Verify that Aspose.Words has access to all required fonts before we render the document to PDF.
        updatedFontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        ASSERT_TRUE(updatedFontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arial"; }));
        ASSERT_TRUE(updatedFontSources[1]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Amethysta"; }));
        ASSERT_TRUE(
            updatedFontSources[1]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Junction Light"; }));

        doc->Save(ArtifactsDir + u"FontSettings.AddFontSource.pdf");

        // Restore the original font sources.
        FontSettings::get_DefaultInstance()->SetFontsSources(originalFontSources);
        //ExEnd
    }

    void SetSpecifyFontFolder()
    {
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->SetFontsFolder(FontsDir, false);

        // Using load options
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_FontSettings(fontSettings);

        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx", loadOptions);

        SharedPtr<FolderFontSource> folderSource = (System::DynamicCast<FolderFontSource>(doc->get_FontSettings()->GetFontsSources()->idx_get(0)));

        ASSERT_EQ(FontsDir, folderSource->get_FolderPath());
        ASSERT_FALSE(folderSource->get_ScanSubfolders());
    }

    void TableSubstitution()
    {
        //ExStart
        //ExFor:Document.FontSettings
        //ExFor:TableSubstitutionRule.SetSubstitutes(String, String[])
        //ExSummary:Shows how set font substitution rules.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(u"Hello world!");
        builder->get_Font()->set_Name(u"Amethysta");
        builder->Writeln(u"The quick brown fox jumps over the lazy dog.");

        ArrayPtr<SharedPtr<FontSourceBase>> fontSources = FontSettings::get_DefaultInstance()->GetFontsSources();

        // The default font sources contain the first font that the document uses.
        ASSERT_EQ(1, fontSources->get_Length());
        ASSERT_TRUE(fontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arial"; }));

        // The second font, "Amethysta", is unavailable.
        ASSERT_FALSE(fontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Amethysta"; }));

        // We can configure a font substitution table which determines
        // which fonts Aspose.Words will use as substitutes for unavailable fonts.
        // Set two substitution fonts for "Amethysta": "Arvo", and "Courier New".
        // If the first substitute is unavailable, Aspose.Words attempts to use the second substitute, and so on.
        doc->set_FontSettings(MakeObject<FontSettings>());
        doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->SetSubstitutes(u"Amethysta",
                                                                                                     MakeArray<String>({u"Arvo", u"Courier New"}));

        // "Amethysta" is unavailable, and the substitution rule states that the first font to use as a substitute is "Arvo".
        ASSERT_FALSE(fontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Arvo"; }));

        // "Arvo" is also unavailable, but "Courier New" is.
        ASSERT_TRUE(fontSources[0]->GetAvailableFonts()->LINQ_Any([](SharedPtr<PhysicalFontInfo> f) { return f->get_FullFontName() == u"Courier New"; }));

        // The output document will display the text that uses the "Amethysta" font formatted with "Courier New".
        doc->Save(ArtifactsDir + u"FontSettings.TableSubstitution.pdf");
        //ExEnd
    }

    void SetSpecifyFontFolders()
    {
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->SetFontsFolders(MakeArray<String>({FontsDir, u"C:\\Windows\\Fonts\\"}), true);

        // Using load options
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_FontSettings(fontSettings);
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx", loadOptions);

        SharedPtr<FolderFontSource> folderSource = (System::DynamicCast<FolderFontSource>(doc->get_FontSettings()->GetFontsSources()->idx_get(0)));
        ASSERT_EQ(FontsDir, folderSource->get_FolderPath());
        ASSERT_TRUE(folderSource->get_ScanSubfolders());

        folderSource = (System::DynamicCast<FolderFontSource>(doc->get_FontSettings()->GetFontsSources()->idx_get(1)));
        ASSERT_EQ(u"C:\\Windows\\Fonts\\", folderSource->get_FolderPath());
        ASSERT_TRUE(folderSource->get_ScanSubfolders());
    }

    void AddFontSubstitutes()
    {
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_SubstitutionSettings()->get_TableSubstitution()->SetSubstitutes(u"Slab", MakeArray<String>({u"Times New Roman", u"Arial"}));
        fontSettings->get_SubstitutionSettings()->get_TableSubstitution()->AddSubstitutes(u"Arvo", MakeArray<String>({u"Open Sans", u"Arial"}));

        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        doc->set_FontSettings(fontSettings);

        ArrayPtr<String> alternativeFonts =
            doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Slab")->LINQ_ToArray();
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Times New Roman", u"Arial"}), alternativeFonts);

        alternativeFonts = doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Arvo")->LINQ_ToArray();
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Open Sans", u"Arial"}), alternativeFonts);
    }

    void FontSourceMemory()
    {
        //ExStart
        //ExFor:Fonts.MemoryFontSource
        //ExFor:Fonts.MemoryFontSource.#ctor(Byte[])
        //ExFor:Fonts.MemoryFontSource.#ctor(Byte[], Int32)
        //ExFor:Fonts.MemoryFontSource.FontData
        //ExFor:Fonts.MemoryFontSource.Type
        //ExSummary:Shows how to use a byte array with data from a font file as a font source.

        ArrayPtr<uint8_t> fontBytes = System::IO::File::ReadAllBytes(MyDir + u"Alte DIN 1451 Mittelschrift.ttf");
        auto memoryFontSource = MakeObject<MemoryFontSource>(fontBytes, 0);

        auto doc = MakeObject<Document>();
        doc->set_FontSettings(MakeObject<FontSettings>());
        doc->get_FontSettings()->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({memoryFontSource}));

        ASSERT_EQ(FontSourceType::MemoryFont, memoryFontSource->get_Type());
        ASSERT_EQ(0, memoryFontSource->get_Priority());
        //ExEnd
    }

    void FontSourceSystem()
    {
        //ExStart
        //ExFor:TableSubstitutionRule.AddSubstitutes(String, String[])
        //ExFor:FontSubstitutionRule.Enabled
        //ExFor:TableSubstitutionRule.GetSubstitutes(String)
        //ExFor:Fonts.FontSettings.ResetFontSources
        //ExFor:Fonts.FontSettings.SubstitutionSettings
        //ExFor:Fonts.FontSubstitutionSettings
        //ExFor:Fonts.SystemFontSource
        //ExFor:Fonts.SystemFontSource.#ctor()
        //ExFor:Fonts.SystemFontSource.#ctor(Int32)
        //ExFor:Fonts.SystemFontSource.GetSystemFontFolders
        //ExFor:Fonts.SystemFontSource.Type
        //ExSummary:Shows how to access a document's system font source and set font substitutes.
        auto doc = MakeObject<Document>();
        doc->set_FontSettings(MakeObject<FontSettings>());

        // By default, a blank document always contains a system font source.
        ASSERT_EQ(1, doc->get_FontSettings()->GetFontsSources()->get_Length());

        auto systemFontSource = System::DynamicCast<SystemFontSource>(doc->get_FontSettings()->GetFontsSources()->idx_get(0));
        ASSERT_EQ(FontSourceType::SystemFonts, systemFontSource->get_Type());
        ASSERT_EQ(0, systemFontSource->get_Priority());

        System::PlatformID pid = System::Environment::get_OSVersion().get_Platform();
        bool isWindows = (pid == System::PlatformID::Win32NT) || (pid == System::PlatformID::Win32S) || (pid == System::PlatformID::Win32Windows) ||
                         (pid == System::PlatformID::WinCE);
        if (isWindows)
        {
            const String fontsPath = u"C:\\WINDOWS\\Fonts";
            ASSERT_EQ(fontsPath.ToLower(), SystemFontSource::GetSystemFontFolders()->LINQ_First().ToLower());
        }

        for (String systemFontFolder : SystemFontSource::GetSystemFontFolders())
        {
            std::cout << systemFontFolder << std::endl;
        }

        // Set a font that exists in the Windows Fonts directory as a substitute for one that does not.
        doc->get_FontSettings()->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(true);
        doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->AddSubstitutes(u"Kreon-Regular", MakeArray<String>({u"Calibri"}));

        ASSERT_EQ(1, doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Kreon-Regular")->LINQ_Count());
        ASSERT_TRUE(doc->get_FontSettings()
                        ->get_SubstitutionSettings()
                        ->get_TableSubstitution()
                        ->GetSubstitutes(u"Kreon-Regular")
                        ->LINQ_ToArray()
                        ->Contains(u"Calibri"));

        // Alternatively, we could add a folder font source in which the corresponding folder contains the font.
        auto folderFontSource = MakeObject<FolderFontSource>(FontsDir, false);
        doc->get_FontSettings()->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({systemFontSource, folderFontSource}));
        ASSERT_EQ(2, doc->get_FontSettings()->GetFontsSources()->get_Length());

        // Resetting the font sources still leaves us with the system font source as well as our substitutes.
        doc->get_FontSettings()->ResetFontSources();

        ASSERT_EQ(1, doc->get_FontSettings()->GetFontsSources()->get_Length());
        ASSERT_EQ(FontSourceType::SystemFonts, doc->get_FontSettings()->GetFontsSources()->idx_get(0)->get_Type());
        ASSERT_EQ(1, doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Kreon-Regular")->LINQ_Count());
        //ExEnd
    }

    void LoadFontFallbackSettingsFromFile()
    {
        //ExStart
        //ExFor:FontFallbackSettings.Load(String)
        //ExFor:FontFallbackSettings.Save(String)
        //ExSummary:Shows how to load and save font fallback settings to/from an XML document in the local file system.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Load an XML document that defines a set of font fallback settings.
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->get_FallbackSettings()->Load(MyDir + u"Font fallback rules.xml");

        doc->set_FontSettings(fontSettings);
        doc->Save(ArtifactsDir + u"FontSettings.LoadFontFallbackSettingsFromFile.pdf");

        // Save our document's current font fallback settings as an XML document.
        doc->get_FontSettings()->get_FallbackSettings()->Save(ArtifactsDir + u"FallbackSettings.xml");
        //ExEnd
    }

    void LoadFontFallbackSettingsFromStream()
    {
        //ExStart
        //ExFor:FontFallbackSettings.Load(Stream)
        //ExFor:FontFallbackSettings.Save(Stream)
        //ExSummary:Shows how to load and save font fallback settings to/from a stream.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Load an XML document that defines a set of font fallback settings.
        {
            auto fontFallbackStream = MakeObject<System::IO::FileStream>(MyDir + u"Font fallback rules.xml", System::IO::FileMode::Open);
            auto fontSettings = MakeObject<FontSettings>();
            fontSettings->get_FallbackSettings()->Load(fontFallbackStream);

            doc->set_FontSettings(fontSettings);
        }

        doc->Save(ArtifactsDir + u"FontSettings.LoadFontFallbackSettingsFromStream.pdf");

        // Use a stream to save our document's current font fallback settings as an XML document.
        {
            auto fontFallbackStream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"FallbackSettings.xml", System::IO::FileMode::Create);
            doc->get_FontSettings()->get_FallbackSettings()->Save(fontFallbackStream);
        }
        //ExEnd

        auto fallbackSettingsDoc = MakeObject<System::Xml::XmlDocument>();
        fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(ArtifactsDir + u"FallbackSettings.xml"));
        auto manager = MakeObject<System::Xml::XmlNamespaceManager>(fallbackSettingsDoc->get_NameTable());
        manager->AddNamespace(u"aw", u"Aspose.Words");

        SharedPtr<System::Xml::XmlNodeList> rules = fallbackSettingsDoc->SelectNodes(u"//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

        ASSERT_EQ(u"0B80-0BFF", rules->idx_get(0)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"Vijaya", rules->idx_get(0)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"1F300-1F64F", rules->idx_get(1)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"Segoe UI Emoji, Segoe UI Symbol", rules->idx_get(1)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"2000-206F, 2070-209F, 20B9", rules->idx_get(2)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"Arial", rules->idx_get(2)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"3040-309F", rules->idx_get(3)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"MS Gothic", rules->idx_get(3)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
        ASSERT_EQ(u"Times New Roman", rules->idx_get(3)->get_Attributes()->idx_get(u"BaseFonts")->get_Value());

        ASSERT_EQ(u"3040-309F", rules->idx_get(4)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"MS Mincho", rules->idx_get(4)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"Arial Unicode MS", rules->idx_get(5)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
    }

    void LoadNotoFontsFallbackSettings()
    {
        //ExStart
        //ExFor:FontFallbackSettings.LoadNotoFallbackSettings
        //ExSummary:Shows how to add predefined font fallback settings for Google Noto fonts.
        auto fontSettings = MakeObject<FontSettings>();

        // These are free fonts licensed under the SIL Open Font License.
        // We can download the fonts here:
        // https://www.google.com/get/noto/#sans-lgc
        fontSettings->SetFontsFolder(FontsDir + u"Noto", false);

        // Note that the predefined settings only use Sans-style Noto fonts with regular weight.
        // Some of the Noto fonts use advanced typography features.
        // Fonts featuring advanced typography may not be rendered correctly as Aspose.Words currently do not support them.
        fontSettings->get_FallbackSettings()->LoadNotoFallbackSettings();
        fontSettings->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(false);
        fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Noto Sans");

        auto doc = MakeObject<Document>();
        doc->set_FontSettings(fontSettings);
        //ExEnd

        TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, u"https://www.google.com/get/noto/#sans-lgc");
    }

    void DefaultFontSubstitutionRule_()
    {
        //ExStart
        //ExFor:Fonts.DefaultFontSubstitutionRule
        //ExFor:Fonts.DefaultFontSubstitutionRule.DefaultFontName
        //ExFor:Fonts.FontSubstitutionSettings.DefaultFontSubstitution
        //ExSummary:Shows how to set the default font substitution rule.
        auto doc = MakeObject<Document>();
        auto fontSettings = MakeObject<FontSettings>();
        doc->set_FontSettings(fontSettings);

        // Get the default substitution rule within FontSettings.
        // This rule will substitute all missing fonts with "Times New Roman".
        SharedPtr<DefaultFontSubstitutionRule> defaultFontSubstitutionRule = fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution();
        ASSERT_TRUE(defaultFontSubstitutionRule->get_Enabled());
        ASSERT_EQ(u"Times New Roman", defaultFontSubstitutionRule->get_DefaultFontName());

        // Set the default font substitute to "Courier New".
        defaultFontSubstitutionRule->set_DefaultFontName(u"Courier New");

        // Using a document builder, add some text in a font that we do not have to see the substitution take place,
        // and then render the result in a PDF.
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Missing Font");
        builder->Writeln(u"Line written in a missing font, which will be substituted with Courier New.");

        doc->Save(ArtifactsDir + u"FontSettings.DefaultFontSubstitutionRule.pdf");
        //ExEnd

        ASSERT_EQ(u"Courier New", defaultFontSubstitutionRule->get_DefaultFontName());
    }

    void FontConfigSubstitution()
    {
        //ExStart
        //ExFor:Fonts.FontConfigSubstitutionRule
        //ExFor:Fonts.FontConfigSubstitutionRule.Enabled
        //ExFor:Fonts.FontConfigSubstitutionRule.IsFontConfigAvailable
        //ExFor:Fonts.FontConfigSubstitutionRule.ResetCache
        //ExFor:Fonts.FontSubstitutionRule
        //ExFor:Fonts.FontSubstitutionRule.Enabled
        //ExFor:Fonts.FontSubstitutionSettings.FontConfigSubstitution
        //ExSummary:Shows operating system-dependent font config substitution.
        auto fontSettings = MakeObject<FontSettings>();
        SharedPtr<FontConfigSubstitutionRule> fontConfigSubstitution = fontSettings->get_SubstitutionSettings()->get_FontConfigSubstitution();

        System::PlatformID pid = System::Environment::get_OSVersion().get_Platform();
        bool isWindows = (pid == System::PlatformID::Win32NT) || (pid == System::PlatformID::Win32S) || (pid == System::PlatformID::Win32Windows) ||
                         (pid == System::PlatformID::WinCE);

        // The FontConfigSubstitutionRule object works differently on Windows/non-Windows platforms.
        // On Windows, it is unavailable.
        if (isWindows)
        {
            ASSERT_FALSE(fontConfigSubstitution->get_Enabled());
            ASSERT_FALSE(fontConfigSubstitution->IsFontConfigAvailable());
        }

        bool isLinuxOrMac = (pid == System::PlatformID::Unix) || (pid == System::PlatformID::MacOSX);

        // On Linux/Mac, we will have access to it, and will be able to perform operations.
        if (isLinuxOrMac)
        {
            ASSERT_TRUE(fontConfigSubstitution->get_Enabled());
            ASSERT_TRUE(fontConfigSubstitution->IsFontConfigAvailable());

            fontConfigSubstitution->ResetCache();
        }
        //ExEnd
    }

    void FallbackSettings()
    {
        //ExStart
        //ExFor:Fonts.FontFallbackSettings.LoadMsOfficeFallbackSettings
        //ExFor:Fonts.FontFallbackSettings.LoadNotoFallbackSettings
        //ExSummary:Shows how to load pre-defined fallback font settings.
        auto doc = MakeObject<Document>();

        auto fontSettings = MakeObject<FontSettings>();
        doc->set_FontSettings(fontSettings);
        SharedPtr<FontFallbackSettings> fontFallbackSettings = fontSettings->get_FallbackSettings();

        // Save the default fallback font scheme to an XML document.
        // For example, one of the elements has a value of "0C00-0C7F" for Range and a corresponding "Vani" value for FallbackFonts.
        // This means that if the font some text is using does not have symbols for the 0x0C00-0x0C7F Unicode block,
        // the fallback scheme will use symbols from the "Vani" font substitute.
        fontFallbackSettings->Save(ArtifactsDir + u"FontSettings.FallbackSettings.Default.xml");

        // Below are two pre-defined font fallback schemes we can choose from.
        // 1 -  Use the default Microsoft Office scheme, which is the same one as the default:
        fontFallbackSettings->LoadMsOfficeFallbackSettings();
        fontFallbackSettings->Save(ArtifactsDir + u"FontSettings.FallbackSettings.LoadMsOfficeFallbackSettings.xml");

        // 2 -  Use the scheme built from Google Noto fonts:
        fontFallbackSettings->LoadNotoFallbackSettings();
        fontFallbackSettings->Save(ArtifactsDir + u"FontSettings.FallbackSettings.LoadNotoFallbackSettings.xml");
        //ExEnd

        auto fallbackSettingsDoc = MakeObject<System::Xml::XmlDocument>();
        fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(ArtifactsDir + u"FontSettings.FallbackSettings.Default.xml"));
        auto manager = MakeObject<System::Xml::XmlNamespaceManager>(fallbackSettingsDoc->get_NameTable());
        manager->AddNamespace(u"aw", u"Aspose.Words");

        SharedPtr<System::Xml::XmlNodeList> rules = fallbackSettingsDoc->SelectNodes(u"//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

        ASSERT_EQ(u"0C00-0C7F", rules->idx_get(5)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"Vani", rules->idx_get(5)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
    }

    void FallbackSettingsCustom()
    {
        //ExStart
        //ExFor:Fonts.FontSettings.FallbackSettings
        //ExFor:Fonts.FontFallbackSettings
        //ExFor:Fonts.FontFallbackSettings.BuildAutomatic
        //ExSummary:Shows how to distribute fallback fonts across Unicode character code ranges.
        auto doc = MakeObject<Document>();

        auto fontSettings = MakeObject<FontSettings>();
        doc->set_FontSettings(fontSettings);
        SharedPtr<FontFallbackSettings> fontFallbackSettings = fontSettings->get_FallbackSettings();

        // Configure our font settings to source fonts only from the "MyFonts" folder.
        auto folderFontSource = MakeObject<FolderFontSource>(FontsDir, false);
        fontSettings->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({folderFontSource}));

        // Calling the "BuildAutomatic" method will generate a fallback scheme that
        // distributes accessible fonts across as many Unicode character codes as possible.
        // In our case, it only has access to the handful of fonts inside the "MyFonts" folder.
        fontFallbackSettings->BuildAutomatic();
        fontFallbackSettings->Save(ArtifactsDir + u"FontSettings.FallbackSettingsCustom.BuildAutomatic.xml");

        // We can also load a custom substitution scheme from a file like this.
        // This scheme applies the "AllegroOpen" font across the "0000-00ff" Unicode blocks, the "AllegroOpen" font across "0100-024f",
        // and the "M+ 2m" font in all other ranges that other fonts in the scheme do not cover.
        fontFallbackSettings->Load(MyDir + u"Custom font fallback settings.xml");

        // Create a document builder and set its font to one that does not exist in any of our sources.
        // Our font settings will invoke the fallback scheme for characters that we type using the unavailable font.
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_Font()->set_Name(u"Missing Font");

        // Use the builder to print every Unicode character from 0x0021 to 0x052F,
        // with descriptive lines dividing Unicode blocks we defined in our custom font fallback scheme.
        for (int i = 0x0021; i < 0x0530; i++)
        {
            switch (i)
            {
            case 0x0021:
                builder->Writeln(u"\n\n0x0021 - 0x00FF: \nBasic Latin/Latin-1 Supplement Unicode blocks in \"AllegroOpen\" font:");
                break;

            case 0x0100:
                builder->Writeln(u"\n\n0x0100 - 0x024F: \nLatin Extended A/B blocks, mostly in \"AllegroOpen\" font:");
                break;

            case 0x0250:
                builder->Writeln(u"\n\n0x0250 - 0x052F: \nIPA/Greek/Cyrillic blocks in \"M+ 2m\" font:");
                break;
            }

            builder->Write(String::Format(u"{0}", System::Convert::ToChar(i)));
        }

        doc->Save(ArtifactsDir + u"FontSettings.FallbackSettingsCustom.pdf");
        //ExEnd

        auto fallbackSettingsDoc = MakeObject<System::Xml::XmlDocument>();
        fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(ArtifactsDir + u"FontSettings.FallbackSettingsCustom.BuildAutomatic.xml"));
        auto manager = MakeObject<System::Xml::XmlNamespaceManager>(fallbackSettingsDoc->get_NameTable());
        manager->AddNamespace(u"aw", u"Aspose.Words");

        SharedPtr<System::Xml::XmlNodeList> rules = fallbackSettingsDoc->SelectNodes(u"//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

        ASSERT_EQ(u"0000-007F", rules->idx_get(0)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"AllegroOpen", rules->idx_get(0)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"0100-017F", rules->idx_get(2)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"AllegroOpen", rules->idx_get(2)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"0250-02AF", rules->idx_get(4)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"M+ 2m", rules->idx_get(4)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());

        ASSERT_EQ(u"0370-03FF", rules->idx_get(7)->get_Attributes()->idx_get(u"Ranges")->get_Value());
        ASSERT_EQ(u"Arvo", rules->idx_get(7)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
    }

    void TableSubstitutionRule_()
    {
        //ExStart
        //ExFor:Fonts.TableSubstitutionRule
        //ExFor:Fonts.TableSubstitutionRule.LoadLinuxSettings
        //ExFor:Fonts.TableSubstitutionRule.LoadWindowsSettings
        //ExFor:Fonts.TableSubstitutionRule.Save(System.IO.Stream)
        //ExFor:Fonts.TableSubstitutionRule.Save(System.String)
        //ExSummary:Shows how to access font substitution tables for Windows and Linux.
        auto doc = MakeObject<Document>();
        auto fontSettings = MakeObject<FontSettings>();
        doc->set_FontSettings(fontSettings);

        // Create a new table substitution rule and load the default Microsoft Windows font substitution table.
        SharedPtr<TableSubstitutionRule> tableSubstitutionRule = fontSettings->get_SubstitutionSettings()->get_TableSubstitution();
        tableSubstitutionRule->LoadWindowsSettings();

        // In Windows, the default substitute for the "Times New Roman CE" font is "Times New Roman".
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Times New Roman"}), tableSubstitutionRule->GetSubstitutes(u"Times New Roman CE")->LINQ_ToArray());

        // We can save the table in the form of an XML document.
        tableSubstitutionRule->Save(ArtifactsDir + u"FontSettings.TableSubstitutionRule.Windows.xml");

        // Linux has its own substitution table.
        // There are multiple substitute fonts for "Times New Roman CE".
        // If the first substitute, "FreeSerif" is also unavailable,
        // this rule will cycle through the others in the array until it finds an available one.
        tableSubstitutionRule->LoadLinuxSettings();
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"FreeSerif", u"Liberation Serif", u"DejaVu Serif"}),
                         tableSubstitutionRule->GetSubstitutes(u"Times New Roman CE")->LINQ_ToArray());

        // Save the Linux substitution table in the form of an XML document using a stream.
        {
            auto fileStream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"FontSettings.TableSubstitutionRule.Linux.xml", System::IO::FileMode::Create);
            tableSubstitutionRule->Save(fileStream);
        }
        //ExEnd

        auto fallbackSettingsDoc = MakeObject<System::Xml::XmlDocument>();
        fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(ArtifactsDir + u"FontSettings.TableSubstitutionRule.Windows.xml"));
        auto manager = MakeObject<System::Xml::XmlNamespaceManager>(fallbackSettingsDoc->get_NameTable());
        manager->AddNamespace(u"aw", u"Aspose.Words");

        SharedPtr<System::Xml::XmlNodeList> rules = fallbackSettingsDoc->SelectNodes(u"//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);

        ASSERT_EQ(u"Times New Roman CE", rules->idx_get(16)->get_Attributes()->idx_get(u"OriginalFont")->get_Value());
        ASSERT_EQ(u"Times New Roman", rules->idx_get(16)->get_Attributes()->idx_get(u"SubstituteFonts")->get_Value());

        fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(ArtifactsDir + u"FontSettings.TableSubstitutionRule.Linux.xml"));
        rules = fallbackSettingsDoc->SelectNodes(u"//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);

        ASSERT_EQ(u"Times New Roman CE", rules->idx_get(31)->get_Attributes()->idx_get(u"OriginalFont")->get_Value());
        ASSERT_EQ(u"FreeSerif, Liberation Serif, DejaVu Serif", rules->idx_get(31)->get_Attributes()->idx_get(u"SubstituteFonts")->get_Value());
    }

    void TableSubstitutionRuleCustom()
    {
        //ExStart
        //ExFor:Fonts.FontSubstitutionSettings.TableSubstitution
        //ExFor:Fonts.TableSubstitutionRule.AddSubstitutes(System.String,System.String[])
        //ExFor:Fonts.TableSubstitutionRule.GetSubstitutes(System.String)
        //ExFor:Fonts.TableSubstitutionRule.Load(System.IO.Stream)
        //ExFor:Fonts.TableSubstitutionRule.Load(System.String)
        //ExFor:Fonts.TableSubstitutionRule.SetSubstitutes(System.String,System.String[])
        //ExSummary:Shows how to work with custom font substitution tables.
        auto doc = MakeObject<Document>();
        auto fontSettings = MakeObject<FontSettings>();
        doc->set_FontSettings(fontSettings);

        // Create a new table substitution rule and load the default Windows font substitution table.
        SharedPtr<TableSubstitutionRule> tableSubstitutionRule = fontSettings->get_SubstitutionSettings()->get_TableSubstitution();

        // If we select fonts exclusively from our folder, we will need a custom substitution table.
        // We will no longer have access to the Microsoft Windows fonts,
        // such as "Arial" or "Times New Roman" since they do not exist in our new font folder.
        auto folderFontSource = MakeObject<FolderFontSource>(FontsDir, false);
        fontSettings->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({folderFontSource}));

        // Below are two ways of loading a substitution table from a file in the local file system.
        // 1 -  From a stream:
        {
            auto fileStream = MakeObject<System::IO::FileStream>(MyDir + u"Font substitution rules.xml", System::IO::FileMode::Open);
            tableSubstitutionRule->Load(fileStream);
        }

        // 2 -  Directly from a file:
        tableSubstitutionRule->Load(MyDir + u"Font substitution rules.xml");

        // Since we no longer have access to "Arial", our font table will first try substitute it with "Nonexistent Font".
        // We do not have this font so that it will move onto the next substitute, "Kreon", found in the "MyFonts" folder.
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Missing Font", u"Kreon"}), tableSubstitutionRule->GetSubstitutes(u"Arial")->LINQ_ToArray());

        // We can expand this table programmatically. We will add an entry that substitutes "Times New Roman" with "Arvo"
        ASSERT_TRUE(tableSubstitutionRule->GetSubstitutes(u"Times New Roman") == nullptr);
        tableSubstitutionRule->AddSubstitutes(u"Times New Roman", MakeArray<String>({u"Arvo"}));
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Arvo"}), tableSubstitutionRule->GetSubstitutes(u"Times New Roman")->LINQ_ToArray());

        // We can add a secondary fallback substitute for an existing font entry with AddSubstitutes().
        // In case "Arvo" is unavailable, our table will look for "M+ 2m" as a second substitute option.
        tableSubstitutionRule->AddSubstitutes(u"Times New Roman", MakeArray<String>({u"M+ 2m"}));
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Arvo", u"M+ 2m"}), tableSubstitutionRule->GetSubstitutes(u"Times New Roman")->LINQ_ToArray());

        // SetSubstitutes() can set a new list of substitute fonts for a font.
        tableSubstitutionRule->SetSubstitutes(u"Times New Roman", MakeArray<String>({u"Squarish Sans CT", u"M+ 2m"}));
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"Squarish Sans CT", u"M+ 2m"}), tableSubstitutionRule->GetSubstitutes(u"Times New Roman")->LINQ_ToArray());

        // Writing text in fonts that we do not have access to will invoke our substitution rules.
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(u"Text written in Arial, to be substituted by Kreon.");

        builder->get_Font()->set_Name(u"Times New Roman");
        builder->Writeln(u"Text written in Times New Roman, to be substituted by Squarish Sans CT.");

        doc->Save(ArtifactsDir + u"FontSettings.TableSubstitutionRule.Custom.pdf");
        //ExEnd
    }

    void ResolveFontsBeforeLoadingDocument()
    {
        //ExStart
        //ExFor:LoadOptions.FontSettings
        //ExSummary:Shows how to designate font substitutes during loading.
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_FontSettings(MakeObject<FontSettings>());

        // Set a font substitution rule for a LoadOptions object.
        // If the document we are loading uses a font which we do not have,
        // this rule will substitute the unavailable font with one that does exist.
        // In this case, all uses of the "MissingFont" will convert to "Comic Sans MS".
        SharedPtr<TableSubstitutionRule> substitutionRule = loadOptions->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution();
        substitutionRule->AddSubstitutes(u"MissingFont", MakeArray<String>({u"Comic Sans MS"}));

        auto doc = MakeObject<Document>(MyDir + u"Missing font.html", loadOptions);

        // At this point such text will still be in "MissingFont".
        // Font substitution will take place when we render the document.
        ASSERT_EQ(u"MissingFont", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Name());

        doc->Save(ArtifactsDir + u"FontSettings.ResolveFontsBeforeLoadingDocument.pdf");
        //ExEnd
    }

    //ExStart
    //ExFor:StreamFontSource
    //ExFor:StreamFontSource.OpenFontDataStream
    //ExSummary:Shows how to load fonts from stream.
    void StreamFontSourceFileRendering()
    {
        auto fontSettings = MakeObject<FontSettings>();
        fontSettings->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({MakeObject<ExFontSettings::StreamFontSourceFile>()}));

        auto builder = MakeObject<DocumentBuilder>();
        builder->get_Document()->set_FontSettings(fontSettings);
        builder->get_Font()->set_Name(u"Kreon-Regular");
        builder->Writeln(u"Test aspose text when saving to PDF.");

        builder->get_Document()->Save(ArtifactsDir + u"FontSettings.StreamFontSourceFileRendering.pdf");
    }

    /// <summary>
    /// Load the font data only when required instead of storing it in the memory for the entire lifetime of the "FontSettings" object.
    /// </summary>
    class StreamFontSourceFile : public StreamFontSource
    {
    public:
        SharedPtr<System::IO::Stream> OpenFontDataStream() override
        {
            return System::IO::File::OpenRead(FontsDir + u"Kreon-Regular.ttf");
        }

    protected:
        virtual ~StreamFontSourceFile()
        {
        }
    };
    //ExEnd
};

} // namespace ApiExamples
