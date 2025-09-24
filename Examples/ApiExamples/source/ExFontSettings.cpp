// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExFontSettings.h"

#include <xml/xml_node_list.h>
#include <xml/xml_node.h>
#include <xml/xml_namespace_manager.h>
#include <xml/xml_name_table.h>
#include <xml/xml_document.h>
#include <xml/xml_attribute_collection.h>
#include <xml/xml_attribute.h>
#include <testing/test_predicates.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/platform_id.h>
#include <system/operating_system.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file.h>
#include <system/func.h>
#include <system/environment.h>
#include <system/enumerator_adapter.h>
#include <system/enum_helpers.h>
#include <system/details/dispose_guard.h>
#include <system/default.h>
#include <system/convert.h>
#include <system/collections/list.h>
#include <system/collections/ilist.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Fonts/TableSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/SystemFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/PhysicalFontInfo.h>
#include <Aspose.Words.Cpp/Model/Fonts/MemoryFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSubstitutionSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceType.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontNameSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontInfoSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontFallbackSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontConfigSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/FileFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/DefaultFontSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Document/WarningType.h>
#include <Aspose.Words.Cpp/Model/Document/WarningSource.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutOptions.h>


using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Loading;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(174703516u, ::Aspose::Words::ApiExamples::ExFontSettings::HandleDocumentWarnings, ThisTypeBaseTypesInfo);

void ExFontSettings::HandleDocumentWarnings::Warning(System::SharedPtr<Aspose::Words::WarningInfo> info)
{
    // We are only interested in fonts being substituted
    if (info->get_WarningType() == Aspose::Words::WarningType::FontSubstitution)
    {
        std::cout << (System::String(u"Font substitution: ") + info->get_Description()) << std::endl;
        FontWarnings->Warning(info);
    }
}

ExFontSettings::HandleDocumentWarnings::HandleDocumentWarnings()
    : FontWarnings(System::MakeObject<Aspose::Words::WarningInfoCollection>())
{
}

RTTI_INFO_IMPL_HASH(1235511231u, ::Aspose::Words::ApiExamples::ExFontSettings::FontSubstitutionWarningCollector, ThisTypeBaseTypesInfo);

void ExFontSettings::FontSubstitutionWarningCollector::Warning(System::SharedPtr<Aspose::Words::WarningInfo> info)
{
    if (info->get_WarningType() == Aspose::Words::WarningType::FontSubstitution)
    {
        FontSubstitutionWarnings->Warning(info);
    }
}

ExFontSettings::FontSubstitutionWarningCollector::FontSubstitutionWarningCollector()
    : FontSubstitutionWarnings(System::MakeObject<Aspose::Words::WarningInfoCollection>())
{
}

RTTI_INFO_IMPL_HASH(86166397u, ::Aspose::Words::ApiExamples::ExFontSettings::FontSourceWarningCollector, ThisTypeBaseTypesInfo);

void ExFontSettings::FontSourceWarningCollector::Warning(System::SharedPtr<Aspose::Words::WarningInfo> info)
{
    FontSubstitutionWarnings->Warning(info);
}

ExFontSettings::FontSourceWarningCollector::FontSourceWarningCollector()
    : FontSubstitutionWarnings(System::MakeObject<Aspose::Words::WarningInfoCollection>())
{
}

RTTI_INFO_IMPL_HASH(4008762824u, ::Aspose::Words::ApiExamples::ExFontSettings::StreamFontSourceFile, ThisTypeBaseTypesInfo);

System::SharedPtr<System::IO::Stream> ExFontSettings::StreamFontSourceFile::OpenFontDataStream()
{
    return System::IO::File::OpenRead(get_FontsDir() + u"Kreon-Regular.ttf");
}

ExFontSettings::StreamFontSourceFile::~StreamFontSourceFile()
{
}


RTTI_INFO_IMPL_HASH(4091418498u, ::Aspose::Words::ApiExamples::ExFontSettings, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExFontSettings : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExFontSettings> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExFontSettings>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExFontSettings> ExFontSettings::s_instance;

} // namespace gtest_test

void ExFontSettings::DefaultFontInstance()
{
    //ExStart
    //ExFor:FontSettings.DefaultInstance
    //ExSummary:Shows how to configure the default font settings instance.
    // Configure the default font settings instance to use the "Courier New" font
    // as a backup substitute when we attempt to use an unknown font.
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Courier New");
    
    ASSERT_TRUE(Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->get_SubstitutionSettings()->get_DefaultFontSubstitution()->get_Enabled());
    
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Non-existent font");
    builder->Write(u"Hello world!");
    
    // This document does not have a FontSettings configuration. When we render the document,
    // the default FontSettings instance will resolve the missing font.
    // Aspose.Words will use "Courier New" to render text that uses the unknown font.
    ASSERT_TRUE(System::TestTools::IsNull(doc->get_FontSettings()));
    
    doc->Save(get_ArtifactsDir() + u"FontSettings.DefaultFontInstance.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, DefaultFontInstance)
{
    s_instance->DefaultFontInstance();
}

} // namespace gtest_test

void ExFontSettings::DefaultFontName()
{
    //ExStart
    //ExFor:DefaultFontSubstitutionRule.DefaultFontName
    //ExSummary:Shows how to specify a default font.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Arial");
    builder->Writeln(u"Hello world!");
    builder->get_Font()->set_Name(u"Arvo");
    builder->Writeln(u"The quick brown fox jumps over the lazy dog.");
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> fontSources = Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->GetFontsSources();
    
    // The font sources that the document uses contain the font "Arial", but not "Arvo".
    ASSERT_EQ(1, fontSources->get_Length());
    ASSERT_TRUE(fontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Arial";
    }))));
    ASSERT_FALSE(fontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Arvo";
    }))));
    
    // Set the "DefaultFontName" property to "Courier New" to,
    // while rendering the document, apply that font in all cases when another font is not available.
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Courier New");
    
    ASSERT_TRUE(fontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Courier New";
    }))));
    
    // Aspose.Words will now use the default font in place of any missing fonts during any rendering calls.
    doc->Save(get_ArtifactsDir() + u"FontSettings.DefaultFontName.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, DefaultFontName)
{
    s_instance->DefaultFontName();
}

} // namespace gtest_test

void ExFontSettings::UpdatePageLayoutWarnings()
{
    // Store the font sources currently used so we can restore them later
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> originalFontSources = Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->GetFontsSources();
    
    // Load the document to render
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
    auto callback = System::MakeObject<Aspose::Words::ApiExamples::ExFontSettings::HandleDocumentWarnings>();
    doc->set_WarningCallback(callback);
    
    // We can choose the default font to use in the case of any missing fonts
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
    
    // For testing we will set Aspose.Words to look for fonts only in a folder which does not exist. Since Aspose.Words won't
    // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default
    // font specified under FontSettings.DefaultFontName. We can pick up on this substitution using our callback
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->SetFontsFolder(System::String::Empty, false);
    
    // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occurred during rendering
    // are stored until the document save and then sent to the appropriate WarningCallback
    doc->UpdatePageLayout();
    
    // Even though the document was rendered previously, any save warnings are notified to the user during document save
    doc->Save(get_ArtifactsDir() + u"FontSettings.UpdatePageLayoutWarnings.pdf");
    
    ASSERT_TRUE(callback->FontWarnings->get_Count() > 0);
    ASSERT_TRUE(callback->FontWarnings->idx_get(0)->get_WarningType() == Aspose::Words::WarningType::FontSubstitution);
    ASSERT_TRUE(callback->FontWarnings->idx_get(0)->get_Description().Contains(u"has not been found"));
    
    // Restore default fonts
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->SetFontsSources(originalFontSources);
}

namespace gtest_test
{

TEST_F(ExFontSettings, UpdatePageLayoutWarnings)
{
    s_instance->UpdatePageLayoutWarnings();
}

} // namespace gtest_test

void ExFontSettings::SubstitutionWarning()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Times New Roman");
    builder->Writeln(u"Hello world!");
    
    auto callback = System::MakeObject<Aspose::Words::ApiExamples::ExFontSettings::FontSubstitutionWarningCollector>();
    doc->set_WarningCallback(callback);
    
    // Store the current collection of font sources, which will be the default font source for every document
    // for which we do not specify a different font source.
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> originalFontSources = Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->GetFontsSources();
    
    // For testing purposes, we will set Aspose.Words to look for fonts only in a folder that does not exist.
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->SetFontsFolder(System::String::Empty, false);
    
    // When rendering the document, there will be no place to find the "Times New Roman" font.
    // This will cause a font substitution warning, which our callback will detect.
    doc->Save(get_ArtifactsDir() + u"FontSettings.SubstitutionWarning.pdf");
    
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->SetFontsSources(originalFontSources);
    
    ASSERT_EQ(1, callback->FontSubstitutionWarnings->get_Count());
    //ExSkip
    ASSERT_TRUE(callback->FontSubstitutionWarnings->idx_get(0)->get_WarningType() == Aspose::Words::WarningType::FontSubstitution);
    ASSERT_TRUE(System::ObjectExt::Equals(callback->FontSubstitutionWarnings->idx_get(0)->get_Description(), u"Font 'Times New Roman' has not been found. Using 'Fanwood' font instead. Reason: first available font."));
}

namespace gtest_test
{

TEST_F(ExFontSettings, SubstitutionWarning)
{
    s_instance->SubstitutionWarning();
}

} // namespace gtest_test

void ExFontSettings::FontSourceWarning()
{
    auto settings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    settings->SetFontsFolder(u"bad folder?", false);
    
    System::SharedPtr<Aspose::Words::Fonts::FontSourceBase> source = settings->GetFontsSources()->idx_get(0);
    auto callback = System::MakeObject<Aspose::Words::ApiExamples::ExFontSettings::FontSourceWarningCollector>();
    source->set_WarningCallback(callback);
    
    // Get the list of fonts to call warning callback.
    System::SharedPtr<System::Collections::Generic::IList<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>>> fontInfos = source->GetAvailableFonts();
    
    ASSERT_TRUE(callback->FontSubstitutionWarnings->idx_get(0)->get_Description().Contains(u"Error loading font from the folder \"bad folder?\""));
}

namespace gtest_test
{

TEST_F(ExFontSettings, FontSourceWarning)
{
    s_instance->FontSourceWarning();
}

} // namespace gtest_test

void ExFontSettings::EnableFontSubstitution()
{
    //ExStart
    //ExFor:FontInfoSubstitutionRule
    //ExFor:FontSubstitutionSettings.FontInfoSubstitution
    //ExFor:LayoutOptions.KeepOriginalFontMetrics
    //ExFor:IWarningCallback
    //ExFor:IWarningCallback.Warning(WarningInfo)
    //ExFor:WarningInfo
    //ExFor:WarningInfo.Description
    //ExFor:WarningInfo.WarningType
    //ExFor:WarningInfoCollection
    //ExFor:WarningInfoCollection.Warning(WarningInfo)
    //ExFor:WarningInfoCollection.Clear
    //ExFor:WarningType
    //ExFor:DocumentBase.WarningCallback
    //ExSummary:Shows how to set the property for finding the closest match for a missing font from the available font sources.
    // Open a document that contains text formatted with a font that does not exist in any of our font sources.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Missing font.docx");
    
    // Assign a callback for handling font substitution warnings.
    auto warningCollector = System::MakeObject<Aspose::Words::WarningInfoCollection>();
    doc->set_WarningCallback(warningCollector);
    
    // Set a default font name and enable font substitution.
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
    fontSettings->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(true);
    
    // Original font metrics should be used after font substitution.
    doc->get_LayoutOptions()->set_KeepOriginalFontMetrics(true);
    
    // We will get a font substitution warning if we save a document with a missing font.
    doc->set_FontSettings(fontSettings);
    doc->Save(get_ArtifactsDir() + u"FontSettings.EnableFontSubstitution.pdf");
    
    for (auto&& info : warningCollector)
    {
        if (info->get_WarningType() == Aspose::Words::WarningType::FontSubstitution)
        {
            std::cout << info->get_Description() << std::endl;
        }
    }
    //ExEnd
    
    // We can also verify warnings in the collection and clear them.
    ASSERT_EQ(Aspose::Words::WarningSource::Layout, warningCollector->idx_get(0)->get_Source());
    ASSERT_EQ(u"Font '28 Days Later' has not been found. Using 'Calibri' font instead. Reason: alternative name from document.", warningCollector->idx_get(0)->get_Description());
    
    warningCollector->Clear();
    
    ASSERT_EQ(0, warningCollector->get_Count());
}

namespace gtest_test
{

TEST_F(ExFontSettings, EnableFontSubstitution)
{
    s_instance->EnableFontSubstitution();
}

} // namespace gtest_test

void ExFontSettings::SubstitutionWarningsClosestMatch()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Bullet points with alternative font.docx");
    
    auto callback = System::MakeObject<Aspose::Words::WarningInfoCollection>();
    doc->set_WarningCallback(callback);
    
    doc->Save(get_ArtifactsDir() + u"FontSettings.SubstitutionWarningsClosestMatch.pdf");
    
    ASSERT_TRUE(System::ObjectExt::Equals(callback->idx_get(0)->get_Description(), u"Font \'SymbolPS\' has not been found. Using \'Wingdings\' font instead. Reason: font info substitution."));
}

namespace gtest_test
{

TEST_F(ExFontSettings, SubstitutionWarningsClosestMatch)
{
    s_instance->SubstitutionWarningsClosestMatch();
}

} // namespace gtest_test

void ExFontSettings::DisableFontSubstitution()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Missing font.docx");
    
    auto callback = System::MakeObject<Aspose::Words::WarningInfoCollection>();
    doc->set_WarningCallback(callback);
    
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
    fontSettings->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(false);
    
    doc->set_FontSettings(fontSettings);
    doc->Save(get_ArtifactsDir() + u"FontSettings.DisableFontSubstitution.pdf");
    
    auto reg = System::MakeObject<System::Text::RegularExpressions::Regex>(u"Font '28 Days Later' has not been found. Using (.*) font instead. Reason: default font setting.");
    
    for (auto&& fontWarning : callback)
    {
        System::SharedPtr<System::Text::RegularExpressions::Match> match = reg->Match(fontWarning->get_Description());
        if (match->get_Success())
        {
            SUCCEED();
            return;
        }
    }
}

namespace gtest_test
{

TEST_F(ExFontSettings, DisableFontSubstitution)
{
    s_instance->DisableFontSubstitution();
}

} // namespace gtest_test

void ExFontSettings::SubstitutionWarnings()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto callback = System::MakeObject<Aspose::Words::WarningInfoCollection>();
    doc->set_WarningCallback(callback);
    
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
    fontSettings->SetFontsFolder(get_FontsDir(), false);
    fontSettings->get_SubstitutionSettings()->get_TableSubstitution()->AddSubstitutes(u"Arial", System::MakeArray<System::String>({u"Arvo", u"Slab"}));
    
    doc->set_FontSettings(fontSettings);
    doc->Save(get_ArtifactsDir() + u"FontSettings.SubstitutionWarnings.pdf");
    
    ASSERT_EQ(u"Font \'Arial\' has not been found. Using \'Arvo\' font instead. Reason: table substitution.", callback->idx_get(0)->get_Description());
    ASSERT_EQ(u"Font \'Times New Roman\' has not been found. Using \'M+ 2m\' font instead. Reason: font info substitution.", callback->idx_get(1)->get_Description());
}

namespace gtest_test
{

TEST_F(ExFontSettings, SkipMono_SubstitutionWarnings)
{
    RecordProperty("category", "SkipMono");
    
    s_instance->SubstitutionWarnings();
}

} // namespace gtest_test

void ExFontSettings::GetSubstitutionWithoutSuffixes()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Get substitution without suffixes.docx");
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> originalFontSources = Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->GetFontsSources();
    
    auto substitutionWarningHandler = System::MakeObject<Aspose::Words::WarningInfoCollection>();
    doc->set_WarningCallback(substitutionWarningHandler);
    
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>>> fontSources = System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>>>(Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->GetFontsSources());
    auto folderFontSource = System::MakeObject<Aspose::Words::Fonts::FolderFontSource>(get_FontsDir(), true);
    fontSources->Add(folderFontSource);
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> updatedFontSources = fontSources->ToArray();
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->SetFontsSources(updatedFontSources);
    
    doc->Save(get_ArtifactsDir() + u"Font.GetSubstitutionWithoutSuffixes.pdf");
    
    ASSERT_EQ(u"Font 'DINOT-Regular' has not been found. Using 'DINOT' font instead. Reason: font name substitution.", substitutionWarningHandler->idx_get(0)->get_Description());
    
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->SetFontsSources(originalFontSources);
}

namespace gtest_test
{

TEST_F(ExFontSettings, GetSubstitutionWithoutSuffixes)
{
    s_instance->GetSubstitutionWithoutSuffixes();
}

} // namespace gtest_test

void ExFontSettings::FontSourceFile()
{
    //ExStart
    //ExFor:FileFontSource
    //ExFor:FileFontSource.#ctor(String)
    //ExFor:FileFontSource.#ctor(String, Int32)
    //ExFor:FileFontSource.FilePath
    //ExFor:FileFontSource.Type
    //ExFor:FontSourceBase
    //ExFor:FontSourceBase.Priority
    //ExFor:FontSourceBase.Type
    //ExFor:FontSourceType
    //ExSummary:Shows how to use a font file in the local file system as a font source.
    auto fileFontSource = System::MakeObject<Aspose::Words::Fonts::FileFontSource>(get_MyDir() + u"Alte DIN 1451 Mittelschrift.ttf", 0);
    
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->set_FontSettings(System::MakeObject<Aspose::Words::Fonts::FontSettings>());
    doc->get_FontSettings()->SetFontsSources(System::MakeArray<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>>({fileFontSource}));
    
    ASSERT_EQ(get_MyDir() + u"Alte DIN 1451 Mittelschrift.ttf", fileFontSource->get_FilePath());
    ASSERT_EQ(Aspose::Words::Fonts::FontSourceType::FontFile, fileFontSource->get_Type());
    ASSERT_EQ(0, fileFontSource->get_Priority());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, FontSourceFile)
{
    s_instance->FontSourceFile();
}

} // namespace gtest_test

void ExFontSettings::FontSourceFolder()
{
    //ExStart
    //ExFor:FolderFontSource
    //ExFor:FolderFontSource.#ctor(String, Boolean)
    //ExFor:FolderFontSource.#ctor(String, Boolean, Int32)
    //ExFor:FolderFontSource.FolderPath
    //ExFor:FolderFontSource.ScanSubfolders
    //ExFor:FolderFontSource.Type
    //ExSummary:Shows how to use a local system folder which contains fonts as a font source.
    
    // Create a font source from a folder that contains font files.
    auto folderFontSource = System::MakeObject<Aspose::Words::Fonts::FolderFontSource>(get_FontsDir(), false, 1);
    
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->set_FontSettings(System::MakeObject<Aspose::Words::Fonts::FontSettings>());
    doc->get_FontSettings()->SetFontsSources(System::MakeArray<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>>({folderFontSource}));
    
    ASSERT_EQ(get_FontsDir(), folderFontSource->get_FolderPath());
    ASPOSE_ASSERT_EQ(false, folderFontSource->get_ScanSubfolders());
    ASSERT_EQ(Aspose::Words::Fonts::FontSourceType::FontsFolder, folderFontSource->get_Type());
    ASSERT_EQ(1, folderFontSource->get_Priority());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, FontSourceFolder)
{
    s_instance->FontSourceFolder();
}

} // namespace gtest_test

void ExFontSettings::SetFontsFolder(bool recursive)
{
    //ExStart
    //ExFor:FontSettings
    //ExFor:FontSettings.SetFontsFolder(String, Boolean)
    //ExSummary:Shows how to set a font source directory.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Arvo");
    builder->Writeln(u"Hello world!");
    builder->get_Font()->set_Name(u"Amethysta");
    builder->Writeln(u"The quick brown fox jumps over the lazy dog.");
    
    // Our font sources do not contain the font that we have used for text in this document.
    // If we use these font settings while rendering this document,
    // Aspose.Words will apply a fallback font to text which has a font that Aspose.Words cannot locate.
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> originalFontSources = Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->GetFontsSources();
    
    ASSERT_EQ(1, originalFontSources->get_Length());
    ASSERT_TRUE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Arial";
    }))));
    
    // The default font sources are missing the two fonts that we are using in this document.
    ASSERT_FALSE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Arvo";
    }))));
    ASSERT_FALSE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Amethysta";
    }))));
    
    // Use the "SetFontsFolder" method to set a directory which will act as a new font source.
    // Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directory
    // that we are passing in the first argument, but not include any fonts in any of that directory's subfolders.
    // Pass "true" as the "recursive" argument to include all font files in the directory that we are passing
    // in the first argument, as well as all the fonts in its subdirectories.
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->SetFontsFolder(get_FontsDir(), recursive);
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> newFontSources = Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->GetFontsSources();
    
    ASSERT_EQ(1, newFontSources->get_Length());
    ASSERT_FALSE(newFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Arial";
    }))));
    ASSERT_TRUE(newFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Arvo";
    }))));
    
    // The "Amethysta" font is in a subfolder of the font directory.
    if (recursive)
    {
        ASSERT_EQ(30, newFontSources[0]->GetAvailableFonts()->get_Count());
        ASSERT_TRUE(newFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
        {
            return f->get_FullFontName() == u"Amethysta";
        }))));
    }
    else
    {
        ASSERT_EQ(18, newFontSources[0]->GetAvailableFonts()->get_Count());
        ASSERT_FALSE(newFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
        {
            return f->get_FullFontName() == u"Amethysta";
        }))));
    }
    
    doc->Save(get_ArtifactsDir() + u"FontSettings.SetFontsFolder.pdf");
    
    // Restore the original font sources.
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->SetFontsSources(originalFontSources);
    //ExEnd
}

namespace gtest_test
{

using ExFontSettings_SetFontsFolder_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExFontSettings::SetFontsFolder)>::type;

struct ExFontSettings_SetFontsFolder : public ExFontSettings, public Aspose::Words::ApiExamples::ExFontSettings, public ::testing::WithParamInterface<ExFontSettings_SetFontsFolder_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExFontSettings_SetFontsFolder, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SetFontsFolder(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExFontSettings_SetFontsFolder, ::testing::ValuesIn(ExFontSettings_SetFontsFolder::TestCases()));

} // namespace gtest_test

void ExFontSettings::SetFontsFolders(bool recursive)
{
    //ExStart
    //ExFor:FontSettings
    //ExFor:FontSettings.SetFontsFolders(String[], Boolean)
    //ExSummary:Shows how to set multiple font source directories.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Amethysta");
    builder->Writeln(u"The quick brown fox jumps over the lazy dog.");
    builder->get_Font()->set_Name(u"Junction Light");
    builder->Writeln(u"The quick brown fox jumps over the lazy dog.");
    
    // Our font sources do not contain the font that we have used for text in this document.
    // If we use these font settings while rendering this document,
    // Aspose.Words will apply a fallback font to text which has a font that Aspose.Words cannot locate.
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> originalFontSources = Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->GetFontsSources();
    
    ASSERT_EQ(1, originalFontSources->get_Length());
    ASSERT_TRUE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Arial";
    }))));
    
    // The default font sources are missing the two fonts that we are using in this document.
    ASSERT_FALSE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Amethysta";
    }))));
    ASSERT_FALSE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Junction Light";
    }))));
    
    // Use the "SetFontsFolders" method to create a font source from each font directory that we pass as the first argument.
    // Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directories
    // that we are passing in the first argument, but not include any fonts from any of the directories' subfolders.
    // Pass "true" as the "recursive" argument to include all font files in the directories that we are passing
    // in the first argument, as well as all the fonts in their subdirectories.
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->SetFontsFolders(System::MakeArray<System::String>({get_FontsDir() + u"/Amethysta", get_FontsDir() + u"/Junction"}), recursive);
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> newFontSources = Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->GetFontsSources();
    
    ASSERT_EQ(2, newFontSources->get_Length());
    ASSERT_FALSE(newFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Arial";
    }))));
    ASSERT_EQ(1, newFontSources[0]->GetAvailableFonts()->get_Count());
    ASSERT_TRUE(newFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Amethysta";
    }))));
    
    // The "Junction" folder itself contains no font files, but has subfolders that do.
    if (recursive)
    {
        ASSERT_EQ(11, newFontSources[1]->GetAvailableFonts()->get_Count());
        ASSERT_TRUE(newFontSources[1]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
        {
            return f->get_FullFontName() == u"Junction Light";
        }))));
    }
    else
    {
        ASSERT_EQ(0, newFontSources[1]->GetAvailableFonts()->get_Count());
    }
    
    doc->Save(get_ArtifactsDir() + u"FontSettings.SetFontsFolders.pdf");
    
    // Restore the original font sources.
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->SetFontsSources(originalFontSources);
    //ExEnd
}

namespace gtest_test
{

using ExFontSettings_SetFontsFolders_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExFontSettings::SetFontsFolders)>::type;

struct ExFontSettings_SetFontsFolders : public ExFontSettings, public Aspose::Words::ApiExamples::ExFontSettings, public ::testing::WithParamInterface<ExFontSettings_SetFontsFolders_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExFontSettings_SetFontsFolders, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SetFontsFolders(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExFontSettings_SetFontsFolders, ::testing::ValuesIn(ExFontSettings_SetFontsFolders::TestCases()));

} // namespace gtest_test

void ExFontSettings::AddFontSource()
{
    //ExStart
    //ExFor:FontSettings
    //ExFor:FontSettings.GetFontsSources()
    //ExFor:FontSettings.SetFontsSources(FontSourceBase[])
    //ExSummary:Shows how to add a font source to our existing font sources.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Arial");
    builder->Writeln(u"Hello world!");
    builder->get_Font()->set_Name(u"Amethysta");
    builder->Writeln(u"The quick brown fox jumps over the lazy dog.");
    builder->get_Font()->set_Name(u"Junction Light");
    builder->Writeln(u"The quick brown fox jumps over the lazy dog.");
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> originalFontSources = Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->GetFontsSources();
    
    ASSERT_EQ(1, originalFontSources->get_Length());
    
    ASSERT_TRUE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Arial";
    }))));
    
    // The default font source is missing two of the fonts that we are using in our document.
    // When we save this document, Aspose.Words will apply fallback fonts to all text formatted with inaccessible fonts.
    ASSERT_FALSE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Amethysta";
    }))));
    ASSERT_FALSE(originalFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Junction Light";
    }))));
    
    // Create a font source from a folder that contains fonts.
    auto folderFontSource = System::MakeObject<Aspose::Words::Fonts::FolderFontSource>(get_FontsDir(), true);
    
    // Apply a new array of font sources that contains the original font sources, as well as our custom fonts.
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> updatedFontSources = System::MakeArray<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>>({originalFontSources[0], folderFontSource});
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->SetFontsSources(updatedFontSources);
    
    // Verify that Aspose.Words has access to all required fonts before we render the document to PDF.
    updatedFontSources = Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->GetFontsSources();
    
    ASSERT_TRUE(updatedFontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Arial";
    }))));
    ASSERT_TRUE(updatedFontSources[1]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Amethysta";
    }))));
    ASSERT_TRUE(updatedFontSources[1]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Junction Light";
    }))));
    
    doc->Save(get_ArtifactsDir() + u"FontSettings.AddFontSource.pdf");
    
    // Restore the original font sources.
    Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->SetFontsSources(originalFontSources);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, AddFontSource)
{
    s_instance->AddFontSource();
}

} // namespace gtest_test

void ExFontSettings::SetSpecifyFontFolder()
{
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    fontSettings->SetFontsFolder(get_FontsDir(), false);
    
    // Using load options
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_FontSettings(fontSettings);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx", loadOptions);
    
    System::SharedPtr<Aspose::Words::Fonts::FolderFontSource> folderSource = (System::ExplicitCast<Aspose::Words::Fonts::FolderFontSource>(doc->get_FontSettings()->GetFontsSources()->idx_get(0)));
    
    ASSERT_EQ(get_FontsDir(), folderSource->get_FolderPath());
    ASSERT_FALSE(folderSource->get_ScanSubfolders());
}

namespace gtest_test
{

TEST_F(ExFontSettings, SetSpecifyFontFolder)
{
    s_instance->SetSpecifyFontFolder();
}

} // namespace gtest_test

void ExFontSettings::TableSubstitution()
{
    //ExStart
    //ExFor:Document.FontSettings
    //ExFor:TableSubstitutionRule.SetSubstitutes(String, String[])
    //ExSummary:Shows how set font substitution rules.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Arial");
    builder->Writeln(u"Hello world!");
    builder->get_Font()->set_Name(u"Amethysta");
    builder->Writeln(u"The quick brown fox jumps over the lazy dog.");
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>> fontSources = Aspose::Words::Fonts::FontSettings::get_DefaultInstance()->GetFontsSources();
    
    // The default font sources contain the first font that the document uses.
    ASSERT_EQ(1, fontSources->get_Length());
    ASSERT_TRUE(fontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Arial";
    }))));
    
    // The second font, "Amethysta", is unavailable.
    ASSERT_FALSE(fontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Amethysta";
    }))));
    
    // We can configure a font substitution table which determines
    // which fonts Aspose.Words will use as substitutes for unavailable fonts.
    // Set two substitution fonts for "Amethysta": "Arvo", and "Courier New".
    // If the first substitute is unavailable, Aspose.Words attempts to use the second substitute, and so on.
    doc->set_FontSettings(System::MakeObject<Aspose::Words::Fonts::FontSettings>());
    doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->SetSubstitutes(u"Amethysta", System::MakeArray<System::String>({u"Arvo", u"Courier New"}));
    
    // "Amethysta" is unavailable, and the substitution rule states that the first font to use as a substitute is "Arvo".
    ASSERT_FALSE(fontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Arvo";
    }))));
    
    // "Arvo" is also unavailable, but "Courier New" is.
    ASSERT_TRUE(fontSources[0]->GetAvailableFonts()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f)>>([](System::SharedPtr<Aspose::Words::Fonts::PhysicalFontInfo> f) -> bool
    {
        return f->get_FullFontName() == u"Courier New";
    }))));
    
    // The output document will display the text that uses the "Amethysta" font formatted with "Courier New".
    doc->Save(get_ArtifactsDir() + u"FontSettings.TableSubstitution.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, TableSubstitution)
{
    s_instance->TableSubstitution();
}

} // namespace gtest_test

void ExFontSettings::SetSpecifyFontFolders()
{
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    fontSettings->SetFontsFolders(System::MakeArray<System::String>({get_FontsDir(), u"C:\\Windows\\Fonts\\"}), true);
    
    // Using load options
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_FontSettings(fontSettings);
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx", loadOptions);
    
    System::SharedPtr<Aspose::Words::Fonts::FolderFontSource> folderSource = (System::ExplicitCast<Aspose::Words::Fonts::FolderFontSource>(doc->get_FontSettings()->GetFontsSources()->idx_get(0)));
    ASSERT_EQ(get_FontsDir(), folderSource->get_FolderPath());
    ASSERT_TRUE(folderSource->get_ScanSubfolders());
    
    folderSource = (System::ExplicitCast<Aspose::Words::Fonts::FolderFontSource>(doc->get_FontSettings()->GetFontsSources()->idx_get(1)));
    ASSERT_EQ(u"C:\\Windows\\Fonts\\", folderSource->get_FolderPath());
    ASSERT_TRUE(folderSource->get_ScanSubfolders());
}

namespace gtest_test
{

TEST_F(ExFontSettings, SetSpecifyFontFolders)
{
    s_instance->SetSpecifyFontFolders();
}

} // namespace gtest_test

void ExFontSettings::AddFontSubstitutes()
{
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    fontSettings->get_SubstitutionSettings()->get_TableSubstitution()->SetSubstitutes(u"Slab", System::MakeArray<System::String>({u"Times New Roman", u"Arial"}));
    fontSettings->get_SubstitutionSettings()->get_TableSubstitution()->AddSubstitutes(u"Arvo", System::MakeArray<System::String>({u"Open Sans", u"Arial"}));
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    doc->set_FontSettings(fontSettings);
    
    System::ArrayPtr<System::String> alternativeFonts = doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Slab")->LINQ_ToArray();
    ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"Times New Roman", u"Arial"}), alternativeFonts);
    
    alternativeFonts = doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Arvo")->LINQ_ToArray();
    ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"Open Sans", u"Arial"}), alternativeFonts);
}

namespace gtest_test
{

TEST_F(ExFontSettings, AddFontSubstitutes)
{
    s_instance->AddFontSubstitutes();
}

} // namespace gtest_test

void ExFontSettings::FontSourceMemory()
{
    //ExStart
    //ExFor:MemoryFontSource
    //ExFor:MemoryFontSource.#ctor(Byte[])
    //ExFor:MemoryFontSource.#ctor(Byte[], Int32)
    //ExFor:MemoryFontSource.FontData
    //ExFor:MemoryFontSource.Type
    //ExSummary:Shows how to use a byte array with data from a font file as a font source.
    
    System::ArrayPtr<uint8_t> fontBytes = System::IO::File::ReadAllBytes(get_MyDir() + u"Alte DIN 1451 Mittelschrift.ttf");
    auto memoryFontSource = System::MakeObject<Aspose::Words::Fonts::MemoryFontSource>(fontBytes, 0);
    
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->set_FontSettings(System::MakeObject<Aspose::Words::Fonts::FontSettings>());
    doc->get_FontSettings()->SetFontsSources(System::MakeArray<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>>({memoryFontSource}));
    
    ASSERT_EQ(Aspose::Words::Fonts::FontSourceType::MemoryFont, memoryFontSource->get_Type());
    ASSERT_EQ(0, memoryFontSource->get_Priority());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, FontSourceMemory)
{
    s_instance->FontSourceMemory();
}

} // namespace gtest_test

void ExFontSettings::FontSourceSystem()
{
    //ExStart
    //ExFor:TableSubstitutionRule.AddSubstitutes(String, String[])
    //ExFor:FontSubstitutionRule.Enabled
    //ExFor:TableSubstitutionRule.GetSubstitutes(String)
    //ExFor:FontSettings.ResetFontSources
    //ExFor:FontSettings.SubstitutionSettings
    //ExFor:FontSubstitutionSettings
    //ExFor:FontSubstitutionSettings.FontNameSubstitution
    //ExFor:SystemFontSource
    //ExFor:SystemFontSource.#ctor
    //ExFor:SystemFontSource.#ctor(Int32)
    //ExFor:SystemFontSource.GetSystemFontFolders
    //ExFor:SystemFontSource.Type
    //ExSummary:Shows how to access a document's system font source and set font substitutes.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->set_FontSettings(System::MakeObject<Aspose::Words::Fonts::FontSettings>());
    
    // By default, a blank document always contains a system font source.
    ASSERT_EQ(1, doc->get_FontSettings()->GetFontsSources()->get_Length());
    
    auto systemFontSource = System::ExplicitCast<Aspose::Words::Fonts::SystemFontSource>(doc->get_FontSettings()->GetFontsSources()->idx_get(0));
    ASSERT_EQ(Aspose::Words::Fonts::FontSourceType::SystemFonts, systemFontSource->get_Type());
    ASSERT_EQ(0, systemFontSource->get_Priority());
    
    System::PlatformID pid = System::Environment::get_OSVersion().get_Platform();
    bool isWindows = (pid == System::PlatformID::Win32NT) || (pid == System::PlatformID::Win32S) || (pid == System::PlatformID::Win32Windows) || (pid == System::PlatformID::WinCE);
    if (isWindows)
    {
        const System::String fontsPath = u"C:\\WINDOWS\\Fonts";
        System::String actual = System::Default<System::String>();
        System::String condExpression = Aspose::Words::Fonts::SystemFontSource::GetSystemFontFolders()->LINQ_FirstOrDefault();
        if (condExpression != nullptr)
        {
            actual = condExpression.ToLower();
        }
        ASSERT_EQ(fontsPath.ToLower(), actual);
    }
    
    for (System::String systemFontFolder : Aspose::Words::Fonts::SystemFontSource::GetSystemFontFolders())
    {
        std::cout << systemFontFolder << std::endl;
    }
    
    
    // Set a font that exists in the Windows Fonts directory as a substitute for one that does not.
    doc->get_FontSettings()->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(true);
    doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->AddSubstitutes(u"Kreon-Regular", System::MakeArray<System::String>({u"Calibri"}));
    
    ASSERT_EQ(1, doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Kreon-Regular")->LINQ_Count());
    ASSERT_TRUE(doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Kreon-Regular")->LINQ_ToArray()->Contains(u"Calibri"));
    
    // Alternatively, we could add a folder font source in which the corresponding folder contains the font.
    auto folderFontSource = System::MakeObject<Aspose::Words::Fonts::FolderFontSource>(get_FontsDir(), false);
    doc->get_FontSettings()->SetFontsSources(System::MakeArray<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>>({systemFontSource, folderFontSource}));
    ASSERT_EQ(2, doc->get_FontSettings()->GetFontsSources()->get_Length());
    
    // Resetting the font sources still leaves us with the system font source as well as our substitutes.
    doc->get_FontSettings()->ResetFontSources();
    
    ASSERT_EQ(1, doc->get_FontSettings()->GetFontsSources()->get_Length());
    ASSERT_EQ(Aspose::Words::Fonts::FontSourceType::SystemFonts, doc->get_FontSettings()->GetFontsSources()->idx_get(0)->get_Type());
    ASSERT_EQ(1, doc->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution()->GetSubstitutes(u"Kreon-Regular")->LINQ_Count());
    ASSERT_TRUE(doc->get_FontSettings()->get_SubstitutionSettings()->get_FontNameSubstitution()->get_Enabled());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, FontSourceSystem)
{
    s_instance->FontSourceSystem();
}

} // namespace gtest_test

void ExFontSettings::LoadFontFallbackSettingsFromFile()
{
    //ExStart
    //ExFor:FontFallbackSettings.Load(String)
    //ExFor:FontFallbackSettings.Save(String)
    //ExSummary:Shows how to load and save font fallback settings to/from an XML document in the local file system.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // Load an XML document that defines a set of font fallback settings.
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    fontSettings->get_FallbackSettings()->Load(get_MyDir() + u"Font fallback rules.xml");
    
    doc->set_FontSettings(fontSettings);
    doc->Save(get_ArtifactsDir() + u"FontSettings.LoadFontFallbackSettingsFromFile.pdf");
    
    // Save our document's current font fallback settings as an XML document.
    doc->get_FontSettings()->get_FallbackSettings()->Save(get_ArtifactsDir() + u"FallbackSettings.xml");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, LoadFontFallbackSettingsFromFile)
{
    s_instance->LoadFontFallbackSettingsFromFile();
}

} // namespace gtest_test

void ExFontSettings::LoadFontFallbackSettingsFromStream()
{
    //ExStart
    //ExFor:FontFallbackSettings.Load(Stream)
    //ExFor:FontFallbackSettings.Save(Stream)
    //ExSummary:Shows how to load and save font fallback settings to/from a stream.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // Load an XML document that defines a set of font fallback settings.
    {
        auto fontFallbackStream = System::MakeObject<System::IO::FileStream>(get_MyDir() + u"Font fallback rules.xml", System::IO::FileMode::Open);
        auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
        fontSettings->get_FallbackSettings()->Load(fontFallbackStream);
        
        doc->set_FontSettings(fontSettings);
    }
    
    doc->Save(get_ArtifactsDir() + u"FontSettings.LoadFontFallbackSettingsFromStream.pdf");
    
    // Use a stream to save our document's current font fallback settings as an XML document.
    {
        auto fontFallbackStream = System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + u"FallbackSettings.xml", System::IO::FileMode::Create);
        doc->get_FontSettings()->get_FallbackSettings()->Save(fontFallbackStream);
    }
    //ExEnd
    
    auto fallbackSettingsDoc = System::MakeObject<System::Xml::XmlDocument>();
    fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(get_ArtifactsDir() + u"FallbackSettings.xml"));
    auto manager = System::MakeObject<System::Xml::XmlNamespaceManager>(fallbackSettingsDoc->get_NameTable());
    manager->AddNamespace(u"aw", u"Aspose.Words");
    
    System::SharedPtr<System::Xml::XmlNodeList> rules = fallbackSettingsDoc->SelectNodes(u"//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);
    
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

namespace gtest_test
{

TEST_F(ExFontSettings, LoadFontFallbackSettingsFromStream)
{
    s_instance->LoadFontFallbackSettingsFromStream();
}

} // namespace gtest_test

void ExFontSettings::LoadNotoFontsFallbackSettings()
{
    //ExStart
    //ExFor:FontFallbackSettings.LoadNotoFallbackSettings
    //ExSummary:Shows how to add predefined font fallback settings for Google Noto fonts.
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    
    // These are free fonts licensed under the SIL Open Font License.
    // We can download the fonts here:
    // https://www.google.com/get/noto/#sans-lgc
    fontSettings->SetFontsFolder(get_FontsDir() + u"Noto", false);
    
    // Note that the predefined settings only use Sans-style Noto fonts with regular weight.
    // Some of the Noto fonts use advanced typography features.
    // Fonts featuring advanced typography may not be rendered correctly as Aspose.Words currently do not support them.
    fontSettings->get_FallbackSettings()->LoadNotoFallbackSettings();
    fontSettings->get_SubstitutionSettings()->get_FontInfoSubstitution()->set_Enabled(false);
    fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Noto Sans");
    
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->set_FontSettings(fontSettings);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, LoadNotoFontsFallbackSettings)
{
    s_instance->LoadNotoFontsFallbackSettings();
}

} // namespace gtest_test

void ExFontSettings::DefaultFontSubstitutionRule()
{
    //ExStart
    //ExFor:DefaultFontSubstitutionRule
    //ExFor:DefaultFontSubstitutionRule.DefaultFontName
    //ExFor:FontSubstitutionSettings.DefaultFontSubstitution
    //ExSummary:Shows how to set the default font substitution rule.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    doc->set_FontSettings(fontSettings);
    
    // Get the default substitution rule within FontSettings.
    // This rule will substitute all missing fonts with "Times New Roman".
    System::SharedPtr<Aspose::Words::Fonts::DefaultFontSubstitutionRule> defaultFontSubstitutionRule = fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution();
    ASSERT_TRUE(defaultFontSubstitutionRule->get_Enabled());
    ASSERT_EQ(u"Times New Roman", defaultFontSubstitutionRule->get_DefaultFontName());
    
    // Set the default font substitute to "Courier New".
    defaultFontSubstitutionRule->set_DefaultFontName(u"Courier New");
    
    // Using a document builder, add some text in a font that we do not have to see the substitution take place,
    // and then render the result in a PDF.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Missing Font");
    builder->Writeln(u"Line written in a missing font, which will be substituted with Courier New.");
    
    doc->Save(get_ArtifactsDir() + u"FontSettings.DefaultFontSubstitutionRule.pdf");
    //ExEnd
    
    ASSERT_EQ(u"Courier New", defaultFontSubstitutionRule->get_DefaultFontName());
}

namespace gtest_test
{

TEST_F(ExFontSettings, DefaultFontSubstitutionRule)
{
    s_instance->DefaultFontSubstitutionRule();
}

} // namespace gtest_test

void ExFontSettings::FontConfigSubstitution()
{
    //ExStart
    //ExFor:FontConfigSubstitutionRule
    //ExFor:FontConfigSubstitutionRule.Enabled
    //ExFor:FontConfigSubstitutionRule.IsFontConfigAvailable
    //ExFor:FontConfigSubstitutionRule.ResetCache
    //ExFor:FontSubstitutionRule
    //ExFor:FontSubstitutionRule.Enabled
    //ExFor:FontSubstitutionSettings.FontConfigSubstitution
    //ExSummary:Shows operating system-dependent font config substitution.
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    System::SharedPtr<Aspose::Words::Fonts::FontConfigSubstitutionRule> fontConfigSubstitution = fontSettings->get_SubstitutionSettings()->get_FontConfigSubstitution();
    
    bool isWindows = System::MakeArray<System::PlatformID>({System::PlatformID::Win32NT, System::PlatformID::Win32S, System::PlatformID::Win32Windows, System::PlatformID::WinCE})->LINQ_Any(static_cast<System::Func<System::PlatformID, bool>>(static_cast<std::function<bool(System::PlatformID p)>>([](System::PlatformID p) -> bool
    {
        return System::Environment::get_OSVersion().get_Platform() == p;
    })));
    
    // The FontConfigSubstitutionRule object works differently on Windows/non-Windows platforms.
    // On Windows, it is unavailable.
    if (isWindows)
    {
        ASSERT_FALSE(fontConfigSubstitution->get_Enabled());
        ASSERT_FALSE(fontConfigSubstitution->IsFontConfigAvailable());
    }
    
    bool isLinuxOrMac = System::MakeArray<System::PlatformID>({System::PlatformID::Unix, System::PlatformID::MacOSX})->LINQ_Any(static_cast<System::Func<System::PlatformID, bool>>(static_cast<std::function<bool(System::PlatformID p)>>([](System::PlatformID p) -> bool
    {
        return System::Environment::get_OSVersion().get_Platform() == p;
    })));
    
    // On Linux/Mac, we will have access to it, and will be able to perform operations.
    if (isLinuxOrMac)
    {
        ASSERT_TRUE(fontConfigSubstitution->get_Enabled());
        ASSERT_TRUE(fontConfigSubstitution->IsFontConfigAvailable());
        
        fontConfigSubstitution->ResetCache();
    }
    
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, FontConfigSubstitution)
{
    s_instance->FontConfigSubstitution();
}

} // namespace gtest_test

void ExFontSettings::FallbackSettings()
{
    //ExStart
    //ExFor:FontFallbackSettings.LoadMsOfficeFallbackSettings
    //ExFor:FontFallbackSettings.LoadNotoFallbackSettings
    //ExSummary:Shows how to load pre-defined fallback font settings.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    doc->set_FontSettings(fontSettings);
    System::SharedPtr<Aspose::Words::Fonts::FontFallbackSettings> fontFallbackSettings = fontSettings->get_FallbackSettings();
    
    // Save the default fallback font scheme to an XML document.
    // For example, one of the elements has a value of "0C00-0C7F" for Range and a corresponding "Vani" value for FallbackFonts.
    // This means that if the font some text is using does not have symbols for the 0x0C00-0x0C7F Unicode block,
    // the fallback scheme will use symbols from the "Vani" font substitute.
    fontFallbackSettings->Save(get_ArtifactsDir() + u"FontSettings.FallbackSettings.Default.xml");
    
    // Below are two pre-defined font fallback schemes we can choose from.
    // 1 -  Use the default Microsoft Office scheme, which is the same one as the default:
    fontFallbackSettings->LoadMsOfficeFallbackSettings();
    fontFallbackSettings->Save(get_ArtifactsDir() + u"FontSettings.FallbackSettings.LoadMsOfficeFallbackSettings.xml");
    
    // 2 -  Use the scheme built from Google Noto fonts:
    fontFallbackSettings->LoadNotoFallbackSettings();
    fontFallbackSettings->Save(get_ArtifactsDir() + u"FontSettings.FallbackSettings.LoadNotoFallbackSettings.xml");
    //ExEnd
    
    auto fallbackSettingsDoc = System::MakeObject<System::Xml::XmlDocument>();
    fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(get_ArtifactsDir() + u"FontSettings.FallbackSettings.Default.xml"));
    auto manager = System::MakeObject<System::Xml::XmlNamespaceManager>(fallbackSettingsDoc->get_NameTable());
    manager->AddNamespace(u"aw", u"Aspose.Words");
    
    System::SharedPtr<System::Xml::XmlNodeList> rules = fallbackSettingsDoc->SelectNodes(u"//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);
    
    ASSERT_EQ(u"0C00-0C7F", rules->idx_get(9)->get_Attributes()->idx_get(u"Ranges")->get_Value());
    ASSERT_EQ(u"Vani", rules->idx_get(9)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
}

namespace gtest_test
{

TEST_F(ExFontSettings, FallbackSettings)
{
    s_instance->FallbackSettings();
}

} // namespace gtest_test

void ExFontSettings::FallbackSettingsCustom()
{
    //ExStart
    //ExFor:FontSettings.FallbackSettings
    //ExFor:FontFallbackSettings
    //ExFor:FontFallbackSettings.BuildAutomatic
    //ExSummary:Shows how to distribute fallback fonts across Unicode character code ranges.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    doc->set_FontSettings(fontSettings);
    System::SharedPtr<Aspose::Words::Fonts::FontFallbackSettings> fontFallbackSettings = fontSettings->get_FallbackSettings();
    
    // Configure our font settings to source fonts only from the "MyFonts" folder.
    auto folderFontSource = System::MakeObject<Aspose::Words::Fonts::FolderFontSource>(get_FontsDir(), false);
    fontSettings->SetFontsSources(System::MakeArray<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>>({folderFontSource}));
    
    // Calling the "BuildAutomatic" method will generate a fallback scheme that
    // distributes accessible fonts across as many Unicode character codes as possible.
    // In our case, it only has access to the handful of fonts inside the "MyFonts" folder.
    fontFallbackSettings->BuildAutomatic();
    fontFallbackSettings->Save(get_ArtifactsDir() + u"FontSettings.FallbackSettingsCustom.BuildAutomatic.xml");
    
    // We can also load a custom substitution scheme from a file like this.
    // This scheme applies the "AllegroOpen" font across the "0000-00ff" Unicode blocks, the "AllegroOpen" font across "0100-024f",
    // and the "M+ 2m" font in all other ranges that other fonts in the scheme do not cover.
    fontFallbackSettings->Load(get_MyDir() + u"Custom font fallback settings.xml");
    
    // Create a document builder and set its font to one that does not exist in any of our sources.
    // Our font settings will invoke the fallback scheme for characters that we type using the unavailable font.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->get_Font()->set_Name(u"Missing Font");
    
    // Use the builder to print every Unicode character from 0x0021 to 0x052F,
    // with descriptive lines dividing Unicode blocks we defined in our custom font fallback scheme.
    for (int32_t i = 0x0021; i < 0x0530; i++)
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
        
        builder->Write(System::String::Format(u"{0}", System::Convert::ToChar(i)));
    }
    
    doc->Save(get_ArtifactsDir() + u"FontSettings.FallbackSettingsCustom.pdf");
    //ExEnd
    
    auto fallbackSettingsDoc = System::MakeObject<System::Xml::XmlDocument>();
    fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(get_ArtifactsDir() + u"FontSettings.FallbackSettingsCustom.BuildAutomatic.xml"));
    auto manager = System::MakeObject<System::Xml::XmlNamespaceManager>(fallbackSettingsDoc->get_NameTable());
    manager->AddNamespace(u"aw", u"Aspose.Words");
    
    System::SharedPtr<System::Xml::XmlNodeList> rules = fallbackSettingsDoc->SelectNodes(u"//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);
    
    ASSERT_EQ(u"0000-007F", rules->idx_get(0)->get_Attributes()->idx_get(u"Ranges")->get_Value());
    ASSERT_EQ(u"AllegroOpen", rules->idx_get(0)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
    
    ASSERT_EQ(u"0100-017F", rules->idx_get(2)->get_Attributes()->idx_get(u"Ranges")->get_Value());
    ASSERT_EQ(u"AllegroOpen", rules->idx_get(2)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
    
    ASSERT_EQ(u"0250-02AF", rules->idx_get(4)->get_Attributes()->idx_get(u"Ranges")->get_Value());
    ASSERT_EQ(u"M+ 2m", rules->idx_get(4)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
    
    ASSERT_EQ(u"0370-03FF", rules->idx_get(7)->get_Attributes()->idx_get(u"Ranges")->get_Value());
    ASSERT_EQ(u"Arvo", rules->idx_get(7)->get_Attributes()->idx_get(u"FallbackFonts")->get_Value());
}

namespace gtest_test
{

TEST_F(ExFontSettings, FallbackSettingsCustom)
{
    s_instance->FallbackSettingsCustom();
}

} // namespace gtest_test

void ExFontSettings::TableSubstitutionRule()
{
    //ExStart
    //ExFor:TableSubstitutionRule
    //ExFor:TableSubstitutionRule.LoadLinuxSettings
    //ExFor:TableSubstitutionRule.LoadWindowsSettings
    //ExFor:TableSubstitutionRule.Save(Stream)
    //ExFor:TableSubstitutionRule.Save(String)
    //ExSummary:Shows how to access font substitution tables for Windows and Linux.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    doc->set_FontSettings(fontSettings);
    
    // Create a new table substitution rule and load the default Microsoft Windows font substitution table.
    System::SharedPtr<Aspose::Words::Fonts::TableSubstitutionRule> tableSubstitutionRule = fontSettings->get_SubstitutionSettings()->get_TableSubstitution();
    tableSubstitutionRule->LoadWindowsSettings();
    
    // In Windows, the default substitute for the "Times New Roman CE" font is "Times New Roman".
    ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"Times New Roman"}), tableSubstitutionRule->GetSubstitutes(u"Times New Roman CE")->LINQ_ToArray());
    
    // We can save the table in the form of an XML document.
    tableSubstitutionRule->Save(get_ArtifactsDir() + u"FontSettings.TableSubstitutionRule.Windows.xml");
    
    // Linux has its own substitution table.
    // There are multiple substitute fonts for "Times New Roman CE".
    // If the first substitute, "FreeSerif" is also unavailable,
    // this rule will cycle through the others in the array until it finds an available one.
    tableSubstitutionRule->LoadLinuxSettings();
    ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"FreeSerif", u"Liberation Serif", u"DejaVu Serif"}), tableSubstitutionRule->GetSubstitutes(u"Times New Roman CE")->LINQ_ToArray());
    
    // Save the Linux substitution table in the form of an XML document using a stream.
    {
        auto fileStream = System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + u"FontSettings.TableSubstitutionRule.Linux.xml", System::IO::FileMode::Create);
        tableSubstitutionRule->Save(fileStream);
    }
    //ExEnd
    
    auto fallbackSettingsDoc = System::MakeObject<System::Xml::XmlDocument>();
    fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(get_ArtifactsDir() + u"FontSettings.TableSubstitutionRule.Windows.xml"));
    auto manager = System::MakeObject<System::Xml::XmlNamespaceManager>(fallbackSettingsDoc->get_NameTable());
    manager->AddNamespace(u"aw", u"Aspose.Words");
    
    System::SharedPtr<System::Xml::XmlNodeList> rules = fallbackSettingsDoc->SelectNodes(u"//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);
    
    ASSERT_EQ(u"Times New Roman CE", rules->idx_get(16)->get_Attributes()->idx_get(u"OriginalFont")->get_Value());
    ASSERT_EQ(u"Times New Roman", rules->idx_get(16)->get_Attributes()->idx_get(u"SubstituteFonts")->get_Value());
    
    fallbackSettingsDoc->LoadXml(System::IO::File::ReadAllText(get_ArtifactsDir() + u"FontSettings.TableSubstitutionRule.Linux.xml"));
    rules = fallbackSettingsDoc->SelectNodes(u"//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);
    
    ASSERT_EQ(u"Times New Roman CE", rules->idx_get(31)->get_Attributes()->idx_get(u"OriginalFont")->get_Value());
    ASSERT_EQ(u"FreeSerif, Liberation Serif, DejaVu Serif", rules->idx_get(31)->get_Attributes()->idx_get(u"SubstituteFonts")->get_Value());
}

namespace gtest_test
{

TEST_F(ExFontSettings, TableSubstitutionRule)
{
    s_instance->TableSubstitutionRule();
}

} // namespace gtest_test

void ExFontSettings::TableSubstitutionRuleCustom()
{
    //ExStart
    //ExFor:FontSubstitutionSettings.TableSubstitution
    //ExFor:TableSubstitutionRule.AddSubstitutes(String,String[])
    //ExFor:TableSubstitutionRule.GetSubstitutes(String)
    //ExFor:TableSubstitutionRule.Load(Stream)
    //ExFor:TableSubstitutionRule.Load(String)
    //ExFor:TableSubstitutionRule.SetSubstitutes(String,String[])
    //ExSummary:Shows how to work with custom font substitution tables.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    doc->set_FontSettings(fontSettings);
    
    // Create a new table substitution rule and load the default Windows font substitution table.
    System::SharedPtr<Aspose::Words::Fonts::TableSubstitutionRule> tableSubstitutionRule = fontSettings->get_SubstitutionSettings()->get_TableSubstitution();
    
    // If we select fonts exclusively from our folder, we will need a custom substitution table.
    // We will no longer have access to the Microsoft Windows fonts,
    // such as "Arial" or "Times New Roman" since they do not exist in our new font folder.
    auto folderFontSource = System::MakeObject<Aspose::Words::Fonts::FolderFontSource>(get_FontsDir(), false);
    fontSettings->SetFontsSources(System::MakeArray<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>>({folderFontSource}));
    
    // Below are two ways of loading a substitution table from a file in the local file system.
    // 1 -  From a stream:
    {
        auto fileStream = System::MakeObject<System::IO::FileStream>(get_MyDir() + u"Font substitution rules.xml", System::IO::FileMode::Open);
        tableSubstitutionRule->Load(fileStream);
    }
    
    // 2 -  Directly from a file:
    tableSubstitutionRule->Load(get_MyDir() + u"Font substitution rules.xml");
    
    // Since we no longer have access to "Arial", our font table will first try substitute it with "Nonexistent Font".
    // We do not have this font so that it will move onto the next substitute, "Kreon", found in the "MyFonts" folder.
    ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"Missing Font", u"Kreon"}), tableSubstitutionRule->GetSubstitutes(u"Arial")->LINQ_ToArray());
    
    // We can expand this table programmatically. We will add an entry that substitutes "Times New Roman" with "Arvo"
    ASSERT_TRUE(System::TestTools::IsNull(tableSubstitutionRule->GetSubstitutes(u"Times New Roman")));
    tableSubstitutionRule->AddSubstitutes(u"Times New Roman", System::MakeArray<System::String>({u"Arvo"}));
    ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"Arvo"}), tableSubstitutionRule->GetSubstitutes(u"Times New Roman")->LINQ_ToArray());
    
    // We can add a secondary fallback substitute for an existing font entry with AddSubstitutes().
    // In case "Arvo" is unavailable, our table will look for "M+ 2m" as a second substitute option.
    tableSubstitutionRule->AddSubstitutes(u"Times New Roman", System::MakeArray<System::String>({u"M+ 2m"}));
    ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"Arvo", u"M+ 2m"}), tableSubstitutionRule->GetSubstitutes(u"Times New Roman")->LINQ_ToArray());
    
    // SetSubstitutes() can set a new list of substitute fonts for a font.
    tableSubstitutionRule->SetSubstitutes(u"Times New Roman", System::MakeArray<System::String>({u"Squarish Sans CT", u"M+ 2m"}));
    ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"Squarish Sans CT", u"M+ 2m"}), tableSubstitutionRule->GetSubstitutes(u"Times New Roman")->LINQ_ToArray());
    
    // Writing text in fonts that we do not have access to will invoke our substitution rules.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->get_Font()->set_Name(u"Arial");
    builder->Writeln(u"Text written in Arial, to be substituted by Kreon.");
    
    builder->get_Font()->set_Name(u"Times New Roman");
    builder->Writeln(u"Text written in Times New Roman, to be substituted by Squarish Sans CT.");
    
    doc->Save(get_ArtifactsDir() + u"FontSettings.TableSubstitutionRule.Custom.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, TableSubstitutionRuleCustom)
{
    s_instance->TableSubstitutionRuleCustom();
}

} // namespace gtest_test

void ExFontSettings::ResolveFontsBeforeLoadingDocument()
{
    //ExStart
    //ExFor:LoadOptions.FontSettings
    //ExSummary:Shows how to designate font substitutes during loading.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_FontSettings(System::MakeObject<Aspose::Words::Fonts::FontSettings>());
    
    // Set a font substitution rule for a LoadOptions object.
    // If the document we are loading uses a font which we do not have,
    // this rule will substitute the unavailable font with one that does exist.
    // In this case, all uses of the "MissingFont" will convert to "Comic Sans MS".
    System::SharedPtr<Aspose::Words::Fonts::TableSubstitutionRule> substitutionRule = loadOptions->get_FontSettings()->get_SubstitutionSettings()->get_TableSubstitution();
    substitutionRule->AddSubstitutes(u"MissingFont", System::MakeArray<System::String>({u"Comic Sans MS"}));
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Missing font.html", loadOptions);
    
    // At this point such text will still be in "MissingFont".
    // Font substitution will take place when we render the document.
    ASSERT_EQ(u"MissingFont", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Name());
    
    doc->Save(get_ArtifactsDir() + u"FontSettings.ResolveFontsBeforeLoadingDocument.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFontSettings, ResolveFontsBeforeLoadingDocument)
{
    s_instance->ResolveFontsBeforeLoadingDocument();
}

} // namespace gtest_test

void ExFontSettings::StreamFontSourceFileRendering()
{
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    fontSettings->SetFontsSources(System::MakeArray<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>>({System::MakeObject<Aspose::Words::ApiExamples::ExFontSettings::StreamFontSourceFile>()}));
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    builder->get_Document()->set_FontSettings(fontSettings);
    builder->get_Font()->set_Name(u"Kreon-Regular");
    builder->Writeln(u"Test aspose text when saving to PDF.");
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"FontSettings.StreamFontSourceFileRendering.pdf");
}

namespace gtest_test
{

TEST_F(ExFontSettings, StreamFontSourceFileRendering)
{
    s_instance->StreamFontSourceFileRendering();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
