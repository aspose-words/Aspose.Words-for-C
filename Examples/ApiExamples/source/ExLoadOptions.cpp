// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExLoadOptions.h"

#include <testing/test_predicates.h>
#include <system/type_info.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/io/directory.h>
#include <system/exceptions.h>
#include <system/enum_helpers.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <drawing/image_converter.h>
#include <drawing/image.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceType.h>
#include <Aspose.Words.Cpp/Model/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Loading/LanguagePreferences.h>
#include <Aspose.Words.Cpp/Model/Loading/EditingLanguage.h>
#include <Aspose.Words.Cpp/Model/Fonts/TableSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSubstitutionSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/WarningType.h>
#include <Aspose.Words.Cpp/Model/Document/WarningSource.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/LoadFormat.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "TestUtil.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Settings;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(729875783u, ::Aspose::Words::ApiExamples::ExLoadOptions::HtmlLinkedResourceLoadingCallback, ThisTypeBaseTypesInfo);

Aspose::Words::Loading::ResourceLoadingAction ExLoadOptions::HtmlLinkedResourceLoadingCallback::ResourceLoading(System::SharedPtr<Aspose::Words::Loading::ResourceLoadingArgs> args)
{
    switch (args->get_ResourceType())
    {
        case Aspose::Words::Loading::ResourceType::CssStyleSheet:
            std::cout << System::String::Format(u"External CSS Stylesheet found upon loading: {0}", args->get_OriginalUri()) << std::endl;
            return Aspose::Words::Loading::ResourceLoadingAction::Default;
        
        case Aspose::Words::Loading::ResourceType::Image:
        {
            std::cout << System::String::Format(u"External Image found upon loading: {0}", args->get_OriginalUri()) << std::endl;
            const System::String newImageFilename = u"Logo.jpg";
            std::cout << System::String::Format(u"\tImage will be substituted with: {0}", newImageFilename) << std::endl;
            System::SharedPtr<System::Drawing::Image> newImage = System::Drawing::Image::FromFile(get_ImageDir() + newImageFilename);
            auto converter = System::MakeObject<System::Drawing::ImageConverter>();
            System::ArrayPtr<uint8_t> imageBytes = System::ExplicitCast<System::Array<uint8_t>>(converter->ConvertTo(newImage, System::ObjectExt::GetType<System::Array<uint8_t>>()));
            args->SetData(imageBytes);
            return Aspose::Words::Loading::ResourceLoadingAction::UserProvided;
        }
        
        default:
            break;
    }
    
    return Aspose::Words::Loading::ResourceLoadingAction::Default;
}

RTTI_INFO_IMPL_HASH(1149712200u, ::Aspose::Words::ApiExamples::ExLoadOptions::DocumentLoadingWarningCallback, ThisTypeBaseTypesInfo);

void ExLoadOptions::DocumentLoadingWarningCallback::Warning(System::SharedPtr<Aspose::Words::WarningInfo> info)
{
    std::cout << System::String::Format(u"Warning: {0}", info->get_WarningType()) << std::endl;
    std::cout << System::String::Format(u"\tSource: {0}", info->get_Source()) << std::endl;
    std::cout << System::String::Format(u"\tDescription: {0}", info->get_Description()) << std::endl;
    mWarnings->Add(info);
}

System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> ExLoadOptions::DocumentLoadingWarningCallback::GetWarnings()
{
    return mWarnings;
}

ExLoadOptions::DocumentLoadingWarningCallback::DocumentLoadingWarningCallback()
    : mWarnings(System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>>())
{
}

RTTI_INFO_IMPL_HASH(2074016742u, ::Aspose::Words::ApiExamples::ExLoadOptions::LoadingProgressCallback, ThisTypeBaseTypesInfo);

const double ExLoadOptions::LoadingProgressCallback::MaxDuration = 0.5;

ExLoadOptions::LoadingProgressCallback::LoadingProgressCallback()
{
    mLoadingStartedAt = System::DateTime::get_Now();
}

void ExLoadOptions::LoadingProgressCallback::Notify(System::SharedPtr<Aspose::Words::Loading::DocumentLoadingArgs> args)
{
    System::DateTime canceledAt = System::DateTime::get_Now();
    double ellapsedSeconds = (canceledAt - mLoadingStartedAt).get_TotalSeconds();
    
    if (ellapsedSeconds > MaxDuration)
    {
        throw System::OperationCanceledException(System::String::Format(u"EstimatedProgress = {0}; CanceledAt = {1}", args->get_EstimatedProgress(), canceledAt));
    }
}


RTTI_INFO_IMPL_HASH(3052477640u, ::Aspose::Words::ApiExamples::ExLoadOptions, ThisTypeBaseTypesInfo);

void ExLoadOptions::TestLoadOptionsWarningCallback(System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> warnings)
{
    ASSERT_EQ(Aspose::Words::WarningType::UnexpectedContent, warnings->idx_get(0)->get_WarningType());
    ASSERT_EQ(Aspose::Words::WarningSource::Docx, warnings->idx_get(0)->get_Source());
    ASSERT_EQ(u"3F01", warnings->idx_get(0)->get_Description());
    
    ASSERT_EQ(Aspose::Words::WarningType::MinorFormattingLoss, warnings->idx_get(1)->get_WarningType());
    ASSERT_EQ(Aspose::Words::WarningSource::Docx, warnings->idx_get(1)->get_Source());
    ASSERT_EQ(u"Import of element 'shapedefaults' is not supported in Docx format by Aspose.Words.", warnings->idx_get(1)->get_Description());
    
    ASSERT_EQ(Aspose::Words::WarningType::MinorFormattingLoss, warnings->idx_get(2)->get_WarningType());
    ASSERT_EQ(Aspose::Words::WarningSource::Docx, warnings->idx_get(2)->get_Source());
    ASSERT_EQ(u"Import of element 'extraClrSchemeLst' is not supported in Docx format by Aspose.Words.", warnings->idx_get(2)->get_Description());
}


namespace gtest_test
{

class ExLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExLoadOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExLoadOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExLoadOptions> ExLoadOptions::s_instance;

} // namespace gtest_test

void ExLoadOptions::LoadOptionsCallback()
{
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_ResourceLoadingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExLoadOptions::HtmlLinkedResourceLoadingCallback>());
    
    // When we load the document, our callback will handle linked resources such as CSS stylesheets and images.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Images.html", loadOptions);
    doc->Save(get_ArtifactsDir() + u"LoadOptions.LoadOptionsCallback.pdf");
}

namespace gtest_test
{

TEST_F(ExLoadOptions, LoadOptionsCallback)
{
    s_instance->LoadOptionsCallback();
}

} // namespace gtest_test

void ExLoadOptions::ConvertShapeToOfficeMath(bool isConvertShapeToOfficeMath)
{
    //ExStart
    //ExFor:LoadOptions.ConvertShapeToOfficeMath
    //ExSummary:Shows how to convert EquationXML shapes to Office Math objects.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    
    // Use this flag to specify whether to convert the shapes with EquationXML attributes
    // to Office Math objects and then load the document.
    loadOptions->set_ConvertShapeToOfficeMath(isConvertShapeToOfficeMath);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Math shapes.docx", loadOptions);
    
    if (isConvertShapeToOfficeMath)
    {
        ASSERT_EQ(16, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
        ASSERT_EQ(34, doc->GetChildNodes(Aspose::Words::NodeType::OfficeMath, true)->get_Count());
    }
    else
    {
        ASSERT_EQ(24, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
        ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::OfficeMath, true)->get_Count());
    }
    //ExEnd
}

namespace gtest_test
{

using ExLoadOptions_ConvertShapeToOfficeMath_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExLoadOptions::ConvertShapeToOfficeMath)>::type;

struct ExLoadOptions_ConvertShapeToOfficeMath : public ExLoadOptions, public Aspose::Words::ApiExamples::ExLoadOptions, public ::testing::WithParamInterface<ExLoadOptions_ConvertShapeToOfficeMath_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExLoadOptions_ConvertShapeToOfficeMath, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ConvertShapeToOfficeMath(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExLoadOptions_ConvertShapeToOfficeMath, ::testing::ValuesIn(ExLoadOptions_ConvertShapeToOfficeMath::TestCases()));

} // namespace gtest_test

void ExLoadOptions::SetEncoding()
{
    //ExStart
    //ExFor:LoadOptions.Encoding
    //ExSummary:Shows how to set the encoding with which to open a document.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_Encoding(System::Text::Encoding::get_ASCII());
    
    // Load the document while passing the LoadOptions object, then verify the document's contents.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"English text.txt", loadOptions);
    
    ASSERT_TRUE(doc->ToString(Aspose::Words::SaveFormat::Text).Contains(u"This is a sample text in English."));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExLoadOptions, SetEncoding)
{
    s_instance->SetEncoding();
}

} // namespace gtest_test

void ExLoadOptions::FontSettings()
{
    //ExStart
    //ExFor:LoadOptions.FontSettings
    //ExSummary:Shows how to apply font substitution settings while loading a document. 
    // Create a FontSettings object that will substitute the "Times New Roman" font
    // with the font "Arvo" from our "MyFonts" folder.
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    fontSettings->SetFontsFolder(get_FontsDir(), false);
    fontSettings->get_SubstitutionSettings()->get_TableSubstitution()->AddSubstitutes(u"Times New Roman", System::MakeArray<System::String>({u"Arvo"}));
    
    // Set that FontSettings object as a property of a newly created LoadOptions object.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_FontSettings(fontSettings);
    
    // Load the document, then render it as a PDF with the font substitution.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx", loadOptions);
    
    doc->Save(get_ArtifactsDir() + u"LoadOptions.FontSettings.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExLoadOptions, FontSettings)
{
    s_instance->FontSettings();
}

} // namespace gtest_test

void ExLoadOptions::LoadOptionsMswVersion()
{
    //ExStart
    //ExFor:LoadOptions.MswVersion
    //ExSummary:Shows how to emulate the loading procedure of a specific Microsoft Word version during document loading.
    // By default, Aspose.Words load documents according to Microsoft Word 2019 specification.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    
    ASSERT_EQ(Aspose::Words::Settings::MsWordVersion::Word2019, loadOptions->get_MswVersion());
    
    // This document is missing the default paragraph formatting style.
    // This default style will be regenerated when we load the document either with Microsoft Word or Aspose.Words.
    loadOptions->set_MswVersion(Aspose::Words::Settings::MsWordVersion::Word2007);
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx", loadOptions);
    
    // The style's line spacing will have this value when loaded by Microsoft Word 2007 specification.
    ASSERT_NEAR(12.95, doc->get_Styles()->get_DefaultParagraphFormat()->get_LineSpacing(), 0.01);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExLoadOptions, LoadOptionsMswVersion)
{
    s_instance->LoadOptionsMswVersion();
}

} // namespace gtest_test

void ExLoadOptions::LoadOptionsWarningCallback()
{
    // Create a new LoadOptions object and set its WarningCallback attribute
    // as an instance of our IWarningCallback implementation.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_WarningCallback(System::MakeObject<Aspose::Words::ApiExamples::ExLoadOptions::DocumentLoadingWarningCallback>());
    
    // Our callback will print all warnings that come up during the load operation.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx", loadOptions);
    
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> warnings = (System::ExplicitCast<Aspose::Words::ApiExamples::ExLoadOptions::DocumentLoadingWarningCallback>(loadOptions->get_WarningCallback()))->GetWarnings();
    ASSERT_EQ(3, warnings->get_Count());
    TestLoadOptionsWarningCallback(warnings);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExLoadOptions, LoadOptionsWarningCallback)
{
    s_instance->LoadOptionsWarningCallback();
}

} // namespace gtest_test

void ExLoadOptions::TempFolder()
{
    //ExStart
    //ExFor:LoadOptions.TempFolder
    //ExSummary:Shows how to use the hard drive instead of memory when loading a document.
    // When we load a document, various elements are temporarily stored in memory as the save operation occurs.
    // We can use this option to use a temporary folder in the local file system instead,
    // which will reduce our application's memory overhead.
    auto options = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    options->set_TempFolder(get_ArtifactsDir() + u"TempFiles");
    
    // The specified temporary folder must exist in the local file system before the load operation.
    System::IO::Directory::CreateDirectory_(options->get_TempFolder());
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx", options);
    
    // The folder will persist with no residual contents from the load operation.
    ASSERT_EQ(0, System::IO::Directory::GetFiles(options->get_TempFolder())->get_Length());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExLoadOptions, TempFolder)
{
    s_instance->TempFolder();
}

} // namespace gtest_test

void ExLoadOptions::AddEditingLanguage()
{
    //ExStart
    //ExFor:LanguagePreferences
    //ExFor:LanguagePreferences.AddEditingLanguage(EditingLanguage)
    //ExFor:LoadOptions.LanguagePreferences
    //ExFor:EditingLanguage
    //ExSummary:Shows how to apply language preferences when loading a document.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->get_LanguagePreferences()->AddEditingLanguage(Aspose::Words::Loading::EditingLanguage::Japanese);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"No default editing language.docx", loadOptions);
    
    int32_t localeIdFarEast = doc->get_Styles()->get_DefaultFont()->get_LocaleIdFarEast();
    std::cout << (localeIdFarEast == (int32_t)Aspose::Words::Loading::EditingLanguage::Japanese ? System::String(u"The document either has no any FarEast language set in defaults or it was set to Japanese originally.") : System::String(u"The document default FarEast language was set to another than Japanese language originally, so it is not overridden.")) << std::endl;
    //ExEnd
    
    ASSERT_EQ((int32_t)Aspose::Words::Loading::EditingLanguage::Japanese, doc->get_Styles()->get_DefaultFont()->get_LocaleIdFarEast());
    
    doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"No default editing language.docx");
    
    ASSERT_EQ((int32_t)Aspose::Words::Loading::EditingLanguage::EnglishUS, doc->get_Styles()->get_DefaultFont()->get_LocaleIdFarEast());
}

namespace gtest_test
{

TEST_F(ExLoadOptions, AddEditingLanguage)
{
    s_instance->AddEditingLanguage();
}

} // namespace gtest_test

void ExLoadOptions::SetEditingLanguageAsDefault()
{
    //ExStart
    //ExFor:LanguagePreferences.DefaultEditingLanguage
    //ExSummary:Shows how set a default language when loading a document.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->get_LanguagePreferences()->set_DefaultEditingLanguage(Aspose::Words::Loading::EditingLanguage::Russian);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"No default editing language.docx", loadOptions);
    
    int32_t localeId = doc->get_Styles()->get_DefaultFont()->get_LocaleId();
    std::cout << (localeId == (int32_t)Aspose::Words::Loading::EditingLanguage::Russian ? System::String(u"The document either has no any language set in defaults or it was set to Russian originally.") : System::String(u"The document default language was set to another than Russian language originally, so it is not overridden.")) << std::endl;
    //ExEnd
    
    ASSERT_EQ((int32_t)Aspose::Words::Loading::EditingLanguage::Russian, doc->get_Styles()->get_DefaultFont()->get_LocaleId());
    
    doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"No default editing language.docx");
    
    ASSERT_EQ((int32_t)Aspose::Words::Loading::EditingLanguage::EnglishUS, doc->get_Styles()->get_DefaultFont()->get_LocaleId());
}

namespace gtest_test
{

TEST_F(ExLoadOptions, SetEditingLanguageAsDefault)
{
    s_instance->SetEditingLanguageAsDefault();
}

} // namespace gtest_test

void ExLoadOptions::ConvertMetafilesToPng()
{
    //ExStart
    //ExFor:LoadOptions.ConvertMetafilesToPng
    //ExSummary:Shows how to convert WMF/EMF to PNG during loading document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto shape = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Image);
    shape->get_ImageData()->SetImage(get_ImageDir() + u"Windows MetaFile.wmf");
    shape->set_Width(100);
    shape->set_Height(100);
    
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Drawing::Shape>>(shape);
    
    doc->Save(get_ArtifactsDir() + u"Image.CreateImageDirectly.docx");
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(1600, 1600, Aspose::Words::Drawing::ImageType::Wmf, shape);
    
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_ConvertMetafilesToPng(true);
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Image.CreateImageDirectly.docx", loadOptions);
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(1666, 1666, Aspose::Words::Drawing::ImageType::Png, shape);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExLoadOptions, ConvertMetafilesToPng)
{
    s_instance->ConvertMetafilesToPng();
}

} // namespace gtest_test

void ExLoadOptions::OpenChmFile()
{
    System::SharedPtr<Aspose::Words::FileFormatInfo> info = Aspose::Words::FileFormatUtil::DetectFileFormat(get_MyDir() + u"HTML help.chm");
    ASSERT_EQ(info->get_LoadFormat(), Aspose::Words::LoadFormat::Chm);
    
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_Encoding(System::Text::Encoding::GetEncoding(u"windows-1251"));
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"HTML help.chm", loadOptions);
}

namespace gtest_test
{

TEST_F(ExLoadOptions, OpenChmFile)
{
    s_instance->OpenChmFile();
}

} // namespace gtest_test

void ExLoadOptions::ProgressCallback()
{
    auto progressCallback = System::MakeObject<Aspose::Words::ApiExamples::ExLoadOptions::LoadingProgressCallback>();
    
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_ProgressCallback(progressCallback);
    
    try
    {
        auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx", loadOptions);
    }
    catch (System::OperationCanceledException& exception)
    {
        std::cout << exception->get_Message() << std::endl;
        
        // Handle loading duration issue.
    }
    
}

namespace gtest_test
{

TEST_F(ExLoadOptions, ProgressCallback)
{
    s_instance->ProgressCallback();
}

} // namespace gtest_test

void ExLoadOptions::IgnoreOleData()
{
    //ExStart
    //ExFor:LoadOptions.IgnoreOleData
    //ExSummary:Shows how to ingore OLE data while loading.
    // Ignoring OLE data may reduce memory consumption and increase performance
    // without data lost in a case when destination format does not support OLE objects.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_IgnoreOleData(true);
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"OLE objects.docx", loadOptions);
    
    doc->Save(get_ArtifactsDir() + u"LoadOptions.IgnoreOleData.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExLoadOptions, IgnoreOleData)
{
    s_instance->IgnoreOleData();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
