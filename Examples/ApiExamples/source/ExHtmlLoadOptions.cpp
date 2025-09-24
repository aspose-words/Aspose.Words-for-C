// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHtmlLoadOptions.h"

#include <testing/test_predicates.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/io/memory_stream.h>
#include <system/enum_helpers.h>
#include <system/date_time.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtListItemCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtListItem.h>
#include <Aspose.Words.Cpp/Model/Loading/HtmlLoadOptions.h>
#include <Aspose.Words.Cpp/Model/Loading/HtmlControlType.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontInfo.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/WarningType.h>
#include <Aspose.Words.Cpp/Model/Document/WarningSource.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/LoadFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/CertificateHolder.h>

#include "TestUtil.h"


using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(262437133u, ::Aspose::Words::ApiExamples::ExHtmlLoadOptions::ListDocumentWarnings, ThisTypeBaseTypesInfo);

void ExHtmlLoadOptions::ListDocumentWarnings::Warning(System::SharedPtr<Aspose::Words::WarningInfo> info)
{
    mWarnings->Add(info);
}

System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> ExHtmlLoadOptions::ListDocumentWarnings::Warnings()
{
    return mWarnings;
}

ExHtmlLoadOptions::ListDocumentWarnings::ListDocumentWarnings()
    : mWarnings(System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>>())
{
}


RTTI_INFO_IMPL_HASH(1082543005u, ::Aspose::Words::ApiExamples::ExHtmlLoadOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExHtmlLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExHtmlLoadOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExHtmlLoadOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExHtmlLoadOptions> ExHtmlLoadOptions::s_instance;

} // namespace gtest_test

void ExHtmlLoadOptions::SupportVml(bool supportVml)
{
    //ExStart
    //ExFor:HtmlLoadOptions
    //ExFor:HtmlLoadOptions.#ctor
    //ExFor:HtmlLoadOptions.SupportVml
    //ExSummary:Shows how to support conditional comments while loading an HTML document.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::HtmlLoadOptions>();
    
    // If the value is true, then we take VML code into account while parsing the loaded document.
    loadOptions->set_SupportVml(supportVml);
    
    // This document contains a JPEG image within "<!--[if gte vml 1]>" tags,
    // and a different PNG image within "<![if !vml]>" tags.
    // If we set the "SupportVml" flag to "true", then Aspose.Words will load the JPEG.
    // If we set this flag to "false", then Aspose.Words will only load the PNG.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"VML conditional.htm", loadOptions);
    
    if (supportVml)
    {
        ASSERT_EQ(Aspose::Words::Drawing::ImageType::Jpeg, (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_ImageData()->get_ImageType());
    }
    else
    {
        ASSERT_EQ(Aspose::Words::Drawing::ImageType::Png, (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_ImageData()->get_ImageType());
    }
    //ExEnd
    
    auto imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    if (supportVml)
    {
        Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    }
    else
    {
        Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Png, imageShape);
    }
}

namespace gtest_test
{

using ExHtmlLoadOptions_SupportVml_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlLoadOptions::SupportVml)>::type;

struct ExHtmlLoadOptions_SupportVml : public ExHtmlLoadOptions, public Aspose::Words::ApiExamples::ExHtmlLoadOptions, public ::testing::WithParamInterface<ExHtmlLoadOptions_SupportVml_Args>
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

TEST_P(ExHtmlLoadOptions_SupportVml, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SupportVml(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlLoadOptions_SupportVml, ::testing::ValuesIn(ExHtmlLoadOptions_SupportVml::TestCases()));

} // namespace gtest_test

void ExHtmlLoadOptions::WebRequestTimeout()
{
    // Create a new HtmlLoadOptions object and verify its timeout threshold for a web request.
    auto options = System::MakeObject<Aspose::Words::Loading::HtmlLoadOptions>();
    
    // When loading an Html document with resources externally linked by a web address URL,
    // Aspose.Words will abort web requests that fail to fetch the resources within this time limit, in milliseconds.
    ASSERT_EQ(100000, options->get_WebRequestTimeout());
    
    // Set a WarningCallback that will record all warnings that occur during loading.
    auto warningCallback = System::MakeObject<Aspose::Words::ApiExamples::ExHtmlLoadOptions::ListDocumentWarnings>();
    options->set_WarningCallback(warningCallback);
    
    // Load such a document and verify that a shape with image data has been created.
    // This linked image will require a web request to load, which will have to complete within our time limit.
    System::String html = System::String::Format(u"\r\n                <html>\r\n                    <img src=\"{0}\" alt=\"Aspose logo\" style=\"width:400px;height:400px;\">\r\n                </html>\r\n            ", get_ImageUrl());
    
    // Set an unreasonable timeout limit and try load the document again.
    options->set_WebRequestTimeout(0);
    auto doc = System::MakeObject<Aspose::Words::Document>(System::MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(html)), options);
    ASSERT_EQ(2, warningCallback->Warnings()->get_Count());
    
    // A web request that fails to obtain an image within the time limit will still produce an image.
    // However, the image will be the red 'x' that commonly signifies missing images.
    auto imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    ASSERT_EQ(924, imageShape->get_ImageData()->get_ImageBytes()->get_Length());
    
    // We can also configure a custom callback to pick up any warnings from timed out web requests.
    ASSERT_EQ(Aspose::Words::WarningSource::Html, warningCallback->Warnings()->idx_get(0)->get_Source());
    ASSERT_EQ(Aspose::Words::WarningType::DataLoss, warningCallback->Warnings()->idx_get(0)->get_WarningType());
    ASSERT_EQ(System::String::Format(u"Couldn't load a resource from \'{0}\'.", get_ImageUrl()), warningCallback->Warnings()->idx_get(0)->get_Description());
    
    ASSERT_EQ(Aspose::Words::WarningSource::Html, warningCallback->Warnings()->idx_get(1)->get_Source());
    ASSERT_EQ(Aspose::Words::WarningType::DataLoss, warningCallback->Warnings()->idx_get(1)->get_WarningType());
    ASSERT_EQ(u"Image has been replaced with a placeholder.", warningCallback->Warnings()->idx_get(1)->get_Description());
    
    doc->Save(get_ArtifactsDir() + u"HtmlLoadOptions.WebRequestTimeout.docx");
}

namespace gtest_test
{

TEST_F(ExHtmlLoadOptions, WebRequestTimeout)
{
    s_instance->WebRequestTimeout();
}

} // namespace gtest_test

void ExHtmlLoadOptions::LoadHtmlFixed()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlFixedSaveOptions>();
    saveOptions->set_SaveFormat(Aspose::Words::SaveFormat::HtmlFixed);
    
    doc->Save(get_ArtifactsDir() + u"HtmlLoadOptions.Fixed.html", saveOptions);
    
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::HtmlLoadOptions>();
    
    auto warningCallback = System::MakeObject<Aspose::Words::ApiExamples::ExHtmlLoadOptions::ListDocumentWarnings>();
    loadOptions->set_WarningCallback(warningCallback);
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"HtmlLoadOptions.Fixed.html", loadOptions);
    ASSERT_EQ(1, warningCallback->Warnings()->get_Count());
    
    ASSERT_EQ(Aspose::Words::WarningSource::Html, warningCallback->Warnings()->idx_get(0)->get_Source());
    ASSERT_EQ(Aspose::Words::WarningType::MajorFormattingLoss, warningCallback->Warnings()->idx_get(0)->get_WarningType());
    ASSERT_EQ(u"The document is fixed-page HTML. Its structure may not be loaded correctly.", warningCallback->Warnings()->idx_get(0)->get_Description());
}

namespace gtest_test
{

TEST_F(ExHtmlLoadOptions, LoadHtmlFixed)
{
    s_instance->LoadHtmlFixed();
}

} // namespace gtest_test

void ExHtmlLoadOptions::EncryptedHtml()
{
    //ExStart
    //ExFor:HtmlLoadOptions.#ctor(String)
    //ExSummary:Shows how to encrypt an Html document, and then open it using a password.
    // Create and sign an encrypted HTML document from an encrypted .docx.
    System::SharedPtr<Aspose::Words::DigitalSignatures::CertificateHolder> certificateHolder = Aspose::Words::DigitalSignatures::CertificateHolder::Create(get_MyDir() + u"morzal.pfx", u"aw");
    
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_Comments(u"Comment");
    signOptions->set_SignTime(System::DateTime::get_Now());
    signOptions->set_DecryptionPassword(u"docPassword");
    
    System::String inputFileName = get_MyDir() + u"Encrypted.docx";
    System::String outputFileName = get_ArtifactsDir() + u"HtmlLoadOptions.EncryptedHtml.html";
    Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(inputFileName, outputFileName, certificateHolder, signOptions);
    
    // To load and read this document, we will need to pass its decryption
    // password using a HtmlLoadOptions object.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::HtmlLoadOptions>(u"docPassword");
    
    ASSERT_EQ(signOptions->get_DecryptionPassword(), loadOptions->get_Password());
    
    auto doc = System::MakeObject<Aspose::Words::Document>(outputFileName, loadOptions);
    
    ASSERT_EQ(u"Test encrypted document.", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHtmlLoadOptions, EncryptedHtml)
{
    s_instance->EncryptedHtml();
}

} // namespace gtest_test

void ExHtmlLoadOptions::BaseUri()
{
    //ExStart
    //ExFor:HtmlLoadOptions.#ctor(LoadFormat,String,String)
    //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
    //ExFor:LoadOptions.LoadFormat
    //ExFor:LoadFormat
    //ExSummary:Shows how to specify a base URI when opening an html document.
    // Suppose we want to load an .html document that contains an image linked by a relative URI
    // while the image is in a different location. In that case, we will need to resolve the relative URI into an absolute one.
    // We can provide a base URI using an HtmlLoadOptions object. 
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::HtmlLoadOptions>(Aspose::Words::LoadFormat::Html, u"", get_ImageDir());
    
    ASSERT_EQ(Aspose::Words::LoadFormat::Html, loadOptions->get_LoadFormat());
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Missing image.html", loadOptions);
    
    // While the image was broken in the input .html, our custom base URI helped us repair the link.
    auto imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->idx_get(0));
    ASSERT_TRUE(imageShape->get_IsImage());
    
    // This output document will display the image that was missing.
    doc->Save(get_ArtifactsDir() + u"HtmlLoadOptions.BaseUri.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"HtmlLoadOptions.BaseUri.docx");
    
    ASSERT_TRUE((System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_ImageData()->get_ImageBytes()->get_Length() > 0);
}

namespace gtest_test
{

TEST_F(ExHtmlLoadOptions, BaseUri)
{
    s_instance->BaseUri();
}

} // namespace gtest_test

void ExHtmlLoadOptions::GetSelectAsSdt()
{
    //ExStart
    //ExFor:HtmlLoadOptions.PreferredControlType
    //ExFor:HtmlControlType
    //ExSummary:Shows how to set preferred type of document nodes that will represent imported <input> and <select> elements.
    const System::String html = u"\r\n                <html>\r\n                    <select name='ComboBox' size='1'>\r\n                        <option value='val1'>item1</option>\r\n                        <option value='val2'></option>\r\n                    </select>\r\n                </html>\r\n            ";
    
    auto htmlLoadOptions = System::MakeObject<Aspose::Words::Loading::HtmlLoadOptions>();
    htmlLoadOptions->set_PreferredControlType(Aspose::Words::Loading::HtmlControlType::StructuredDocumentTag);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(System::MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(html)), htmlLoadOptions);
    System::SharedPtr<Aspose::Words::NodeCollection> nodes = doc->GetChildNodes(Aspose::Words::NodeType::StructuredDocumentTag, true);
    
    auto tag = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(nodes->idx_get(0));
    //ExEnd
    
    ASSERT_EQ(2, tag->get_ListItems()->get_Count());
    
    ASSERT_EQ(u"val1", tag->get_ListItems()->idx_get(0)->get_Value());
    ASSERT_EQ(u"val2", tag->get_ListItems()->idx_get(1)->get_Value());
}

namespace gtest_test
{

TEST_F(ExHtmlLoadOptions, GetSelectAsSdt)
{
    s_instance->GetSelectAsSdt();
}

} // namespace gtest_test

void ExHtmlLoadOptions::GetInputAsFormField()
{
    const System::String html = u"\r\n                <html>\r\n                    <input type='text' value='Input value text' />\r\n                </html>\r\n            ";
    
    // By default, "HtmlLoadOptions.PreferredControlType" value is "HtmlControlType.FormField".
    // So, we do not set this value.
    auto htmlLoadOptions = System::MakeObject<Aspose::Words::Loading::HtmlLoadOptions>();
    
    auto doc = System::MakeObject<Aspose::Words::Document>(System::MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(html)), htmlLoadOptions);
    System::SharedPtr<Aspose::Words::NodeCollection> nodes = doc->GetChildNodes(Aspose::Words::NodeType::FormField, true);
    
    ASSERT_EQ(1, nodes->get_Count());
    
    auto formField = System::ExplicitCast<Aspose::Words::Fields::FormField>(nodes->idx_get(0));
    ASSERT_EQ(u"Input value text", formField->get_Result());
}

namespace gtest_test
{

TEST_F(ExHtmlLoadOptions, GetInputAsFormField)
{
    s_instance->GetInputAsFormField();
}

} // namespace gtest_test

void ExHtmlLoadOptions::IgnoreNoscriptElements(bool ignoreNoscriptElements)
{
    //ExStart
    //ExFor:HtmlLoadOptions.IgnoreNoscriptElements
    //ExSummary:Shows how to ignore <noscript> HTML elements.
    const System::String html = u"\r\n                <html>\r\n                  <head>\r\n                    <title>NOSCRIPT</title>\r\n                      <meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\n                      <script type=\"text/javascript\">\r\n                        alert(\"Hello, world!\");\r\n                      </script>\r\n                  </head>\r\n                <body>\r\n                  <noscript><p>Your browser does not support JavaScript!</p></noscript>\r\n                </body>\r\n                </html>";
    
    auto htmlLoadOptions = System::MakeObject<Aspose::Words::Loading::HtmlLoadOptions>();
    htmlLoadOptions->set_IgnoreNoscriptElements(ignoreNoscriptElements);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(System::MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(html)), htmlLoadOptions);
    doc->Save(get_ArtifactsDir() + u"HtmlLoadOptions.IgnoreNoscriptElements.pdf");
    //ExEnd
}

namespace gtest_test
{

using ExHtmlLoadOptions_IgnoreNoscriptElements_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlLoadOptions::IgnoreNoscriptElements)>::type;

struct ExHtmlLoadOptions_IgnoreNoscriptElements : public ExHtmlLoadOptions, public Aspose::Words::ApiExamples::ExHtmlLoadOptions, public ::testing::WithParamInterface<ExHtmlLoadOptions_IgnoreNoscriptElements_Args>
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

TEST_P(ExHtmlLoadOptions_IgnoreNoscriptElements, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreNoscriptElements(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlLoadOptions_IgnoreNoscriptElements, ::testing::ValuesIn(ExHtmlLoadOptions_IgnoreNoscriptElements::TestCases()));

} // namespace gtest_test

void ExHtmlLoadOptions::BlockImport(Aspose::Words::Loading::BlockImportMode blockImportMode)
{
    //ExStart
    //ExFor:HtmlLoadOptions.BlockImportMode
    //ExFor:BlockImportMode
    //ExSummary:Shows how properties of block-level elements are imported from HTML-based documents.
    const System::String html = u"\r\n            <html>\r\n                <div style='border:dotted'>\r\n                    <div style='border:solid'>\r\n                        <p>paragraph 1</p>\r\n                        <p>paragraph 2</p>\r\n                    </div>\r\n                </div>\r\n            </html>";
    auto stream = System::MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(html));
    
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::HtmlLoadOptions>();
    // Set the new mode of import HTML block-level elements.
    loadOptions->set_BlockImportMode(blockImportMode);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(stream, loadOptions);
    doc->Save(get_ArtifactsDir() + u"HtmlLoadOptions.BlockImport.docx");
    //ExEnd
}

namespace gtest_test
{

using ExHtmlLoadOptions_BlockImport_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlLoadOptions::BlockImport)>::type;

struct ExHtmlLoadOptions_BlockImport : public ExHtmlLoadOptions, public Aspose::Words::ApiExamples::ExHtmlLoadOptions, public ::testing::WithParamInterface<ExHtmlLoadOptions_BlockImport_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Loading::BlockImportMode::Preserve),
            std::make_tuple(Aspose::Words::Loading::BlockImportMode::Merge),
        };
    }
};

TEST_P(ExHtmlLoadOptions_BlockImport, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->BlockImport(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlLoadOptions_BlockImport, ::testing::ValuesIn(ExHtmlLoadOptions_BlockImport::TestCases()));

} // namespace gtest_test

void ExHtmlLoadOptions::FontFaceRules()
{
    //ExStart:FontFaceRules
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:HtmlLoadOptions.SupportFontFaceRules
    //ExSummary:Shows how to load declared "@font-face" rules.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::HtmlLoadOptions>();
    loadOptions->set_SupportFontFaceRules(true);
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Html with FontFace.html", loadOptions);
    
    ASSERT_EQ(u"Squarish Sans CT Regular", doc->get_FontInfos()->idx_get(0)->get_Name());
    //ExEnd:FontFaceRules
}

namespace gtest_test
{

TEST_F(ExHtmlLoadOptions, FontFaceRules)
{
    s_instance->FontFaceRules();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
