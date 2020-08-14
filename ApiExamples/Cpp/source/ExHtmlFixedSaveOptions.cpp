// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHtmlFixedSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/path.h>
#include <system/io/memory_stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file_info.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/console.h>
#include <gtest/gtest.h>
#include <functional>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/ResourceSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IResourceSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlFixedPageHorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Saving/ExportFontFormat.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1435199370u, ::ApiExamples::ExHtmlFixedSaveOptions::ResourceUriPrinter, ThisTypeBaseTypesInfo);

void ExHtmlFixedSaveOptions::ResourceUriPrinter::ResourceSaving(SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args)
{
    // If we set a folder alias in the SaveOptions object, it will be printed here
    System::Console::WriteLine(String::Format(u"Resource #{0} \"{1}\"",++mSavedResourceCount,args->get_ResourceFileName()));

    String extension = System::IO::Path::GetExtension(args->get_ResourceFileName());
    const String& switch_value_1 = extension;
    if (switch_value_1 == u".ttf" || switch_value_1 == u".woff")
    {

        {
            // By default 'ResourceFileUri' used system folder for fonts
            // To avoid problems across platforms you must explicitly specify the path for the fonts
            args->set_ResourceFileUri(ArtifactsDir + System::IO::Path::DirectorySeparatorChar + args->get_ResourceFileName());
            goto switch_break_1;
        }
    }
    switch_break_1:;
    System::Console::WriteLine(String(u"\t") + args->get_ResourceFileUri());

    // If we specified a ResourcesFolderAlias we will also need to redirect each stream to put its resource in that folder
    args->set_ResourceStream(MakeObject<System::IO::FileStream>(args->get_ResourceFileUri(), System::IO::FileMode::Create));
    args->set_KeepResourceStreamOpen(false);
}

ExHtmlFixedSaveOptions::ResourceUriPrinter::ResourceUriPrinter() : mSavedResourceCount(0)
{
}

RTTI_INFO_IMPL_HASH(640691227u, ::ApiExamples::ExHtmlFixedSaveOptions::ResourceSavingCallback, ThisTypeBaseTypesInfo);

void ExHtmlFixedSaveOptions::ResourceSavingCallback::ResourceSaving(SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args)
{
    System::Console::WriteLine(String::Format(u"Original document URI:\t{0}",args->get_Document()->get_OriginalFileName()));
    System::Console::WriteLine(String::Format(u"Resource being saved:\t{0}",args->get_ResourceFileName()));
    System::Console::WriteLine(String::Format(u"Full uri after saving:\t{0}",args->get_ResourceFileUri()));

    args->set_ResourceStream(MakeObject<System::IO::MemoryStream>());
    args->set_KeepResourceStreamOpen(true);

    String extension = System::IO::Path::GetExtension(args->get_ResourceFileName());
    const String& switch_value_0 = extension;
    if (switch_value_0 == u".ttf" || switch_value_0 == u".woff")
    {

        {
            FAIL() << "'ResourceSavingCallback' is not fired for fonts when 'UseTargetMachineFonts' is true";
            goto switch_break_0;
        }
    }
    switch_break_0:;
}

namespace gtest_test
{

class ExHtmlFixedSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExHtmlFixedSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExHtmlFixedSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExHtmlFixedSaveOptions> ExHtmlFixedSaveOptions::s_instance;

} // namespace gtest_test

void ExHtmlFixedSaveOptions::UseEncoding()
{
    //ExStart
    //ExFor:HtmlFixedSaveOptions.Encoding
    //ExSummary:Shows how to set encoding while exporting to HTML.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Writeln(u"Hello World!");

    // The default encoding is UTF-8
    // If we want to represent our document using a different encoding, we can set one explicitly using a SaveOptions object
    auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
    htmlFixedSaveOptions->set_Encoding(System::Text::Encoding::GetEncoding(u"ASCII"));

    ASSERT_EQ(u"US-ASCII", htmlFixedSaveOptions->get_Encoding()->get_EncodingName());

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.UseEncoding.html", htmlFixedSaveOptions);
    //ExEnd

    ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.UseEncoding.html"), u"content=\"text/html; charset=us-ascii\"")->get_Success());
}

namespace gtest_test
{

TEST_F(ExHtmlFixedSaveOptions, UseEncoding)
{
    s_instance->UseEncoding();
}

} // namespace gtest_test

void ExHtmlFixedSaveOptions::GetEncoding()
{
    SharedPtr<Document> doc = DocumentHelper::CreateDocumentFillWithDummyText();

    auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
    htmlFixedSaveOptions->set_Encoding(System::Text::Encoding::GetEncoding(u"utf-16"));

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.GetEncoding.html", htmlFixedSaveOptions);
}

namespace gtest_test
{

TEST_F(ExHtmlFixedSaveOptions, GetEncoding)
{
    s_instance->GetEncoding();
}

} // namespace gtest_test

void ExHtmlFixedSaveOptions::ExportEmbeddedCSS(bool doExportEmbeddedCss)
{
    //ExStart
    //ExFor:HtmlFixedSaveOptions.ExportEmbeddedCss
    //ExSummary:Shows how to export embedded stylesheets into an HTML file.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
    htmlFixedSaveOptions->set_ExportEmbeddedCss(doExportEmbeddedCss);

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedCSS.html", htmlFixedSaveOptions);

    String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedCSS.html");

    if (doExportEmbeddedCss)
    {
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"<style type=\"text/css\">")->get_Success());
        ASSERT_FALSE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedCSS/styles.css"));
    }
    else
    {
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"<link rel=\"stylesheet\" type=\"text/css\" href=\"HtmlFixedSaveOptions[.]ExportEmbeddedCSS/styles[.]css\" media=\"all\" />")->get_Success());
        ASSERT_TRUE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedCSS/styles.css"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExportEmbeddedCSS_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportEmbeddedCSS)>::type;

struct ExportEmbeddedCSSVP : public ExHtmlFixedSaveOptions, public ApiExamples::ExHtmlFixedSaveOptions, public ::testing::WithParamInterface<ExportEmbeddedCSS_Args>
{
    static std::vector<ExportEmbeddedCSS_Args> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExportEmbeddedCSSVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportEmbeddedCSS(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlFixedSaveOptions, ExportEmbeddedCSSVP, ::testing::ValuesIn(ExportEmbeddedCSSVP::TestCases()));

} // namespace gtest_test

void ExHtmlFixedSaveOptions::ExportEmbeddedFonts(bool doExportEmbeddedFonts)
{
    //ExStart
    //ExFor:HtmlFixedSaveOptions.ExportEmbeddedFonts
    //ExSummary:Shows how to export embedded fonts into an HTML file.
    auto doc = MakeObject<Document>(MyDir + u"Embedded font.docx");

    auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
    htmlFixedSaveOptions->set_ExportEmbeddedFonts(doExportEmbeddedFonts);

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedFonts.html", htmlFixedSaveOptions);

    String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedFonts/styles.css");

    if (doExportEmbeddedFonts)
    {
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(].+[)] format[(]'woff'[)]; }")->get_Success());

        ASSERT_EQ(0, System::IO::Directory::GetFiles(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedFonts")->LINQ_Count([](String f) { return f.EndsWith(u".woff"); }));
    }
    else
    {
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(]'font001[.]woff'[)] format[(]'woff'[)]; }")->get_Success());

        ASSERT_EQ(2, System::IO::Directory::GetFiles(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedFonts")->LINQ_Count([](String f) { return f.EndsWith(u".woff"); }));
    }
    //ExEnd
}

namespace gtest_test
{

using ExportEmbeddedFonts_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportEmbeddedFonts)>::type;

struct ExportEmbeddedFontsVP : public ExHtmlFixedSaveOptions, public ApiExamples::ExHtmlFixedSaveOptions, public ::testing::WithParamInterface<ExportEmbeddedFonts_Args>
{
    static std::vector<ExportEmbeddedFonts_Args> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExportEmbeddedFontsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportEmbeddedFonts(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlFixedSaveOptions, ExportEmbeddedFontsVP, ::testing::ValuesIn(ExportEmbeddedFontsVP::TestCases()));

} // namespace gtest_test

void ExHtmlFixedSaveOptions::ExportEmbeddedImages(bool doExportImages)
{
    //ExStart
    //ExFor:HtmlFixedSaveOptions.ExportEmbeddedImages
    //ExSummary:Shows how to export embedded images into an HTML file.
    auto doc = MakeObject<Document>(MyDir + u"Images.docx");

    auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
    htmlFixedSaveOptions->set_ExportEmbeddedImages(doExportImages);

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedImages.html", htmlFixedSaveOptions);

    String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedImages.html");

    if (doExportImages)
    {
        ASSERT_FALSE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedImages/image001.jpeg"));
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"<img class=\"awimg\" style=\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\" src=\".+\" />")->get_Success());
    }
    else
    {
        ASSERT_TRUE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedImages/image001.jpeg"));
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, String(u"<img class=\"awimg\" style=\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\" ") + u"src=\"HtmlFixedSaveOptions[.]ExportEmbeddedImages/image001[.]jpeg\" />")->get_Success());
    }
    //ExEnd
}

namespace gtest_test
{

using ExportEmbeddedImages_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportEmbeddedImages)>::type;

struct ExportEmbeddedImagesVP : public ExHtmlFixedSaveOptions, public ApiExamples::ExHtmlFixedSaveOptions, public ::testing::WithParamInterface<ExportEmbeddedImages_Args>
{
    static std::vector<ExportEmbeddedImages_Args> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExportEmbeddedImagesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportEmbeddedImages(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlFixedSaveOptions, ExportEmbeddedImagesVP, ::testing::ValuesIn(ExportEmbeddedImagesVP::TestCases()));

} // namespace gtest_test

void ExHtmlFixedSaveOptions::ExportEmbeddedSvgs(bool doExportSvgs)
{
    //ExStart
    //ExFor:HtmlFixedSaveOptions.ExportEmbeddedSvg
    //ExSummary:Shows how to export embedded SVG objects into an HTML file.
    auto doc = MakeObject<Document>(MyDir + u"Images.docx");

    auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
    htmlFixedSaveOptions->set_ExportEmbeddedSvg(doExportSvgs);

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedSvgs.html", htmlFixedSaveOptions);

    String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedSvgs.html");

    if (doExportSvgs)
    {
        ASSERT_FALSE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001.svg"));
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"<image id=\"image004\" xlink:href=.+/>")->get_Success());
    }
    else
    {
        ASSERT_TRUE(System::IO::File::Exists(ArtifactsDir + u"HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001.svg"));
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"<object type=\"image/svg[+]xml\" data=\"HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001[.]svg\"></object>")->get_Success());
    }
    //ExEnd
}

namespace gtest_test
{

using ExportEmbeddedSvgs_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportEmbeddedSvgs)>::type;

struct ExportEmbeddedSvgsVP : public ExHtmlFixedSaveOptions, public ApiExamples::ExHtmlFixedSaveOptions, public ::testing::WithParamInterface<ExportEmbeddedSvgs_Args>
{
    static std::vector<ExportEmbeddedSvgs_Args> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExportEmbeddedSvgsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportEmbeddedSvgs(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlFixedSaveOptions, ExportEmbeddedSvgsVP, ::testing::ValuesIn(ExportEmbeddedSvgsVP::TestCases()));

} // namespace gtest_test

void ExHtmlFixedSaveOptions::ExportFormFields(bool doExportFormFields)
{
    //ExStart
    //ExFor:HtmlFixedSaveOptions.ExportFormFields
    //ExSummary:Show how to exporting form fields from a document into HTML file.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->InsertCheckBox(u"CheckBox", false, 15);

    auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
    htmlFixedSaveOptions->set_ExportFormFields(doExportFormFields);

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.ExportFormFields.html", htmlFixedSaveOptions);

    String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.ExportFormFields.html");

    if (doExportFormFields)
    {
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, String(u"<a name=\"CheckBox\" style=\"left:0pt; top:0pt;\"></a>") + u"<input style=\"position:absolute; left:0pt; top:0pt;\" type=\"checkbox\" name=\"CheckBox\" />")->get_Success());
    }
    else
    {
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, String(u"<a name=\"CheckBox\" style=\"left:0pt; top:0pt;\"></a>") + u"<div class=\"awdiv\" style=\"left:0.8pt; top:0.8pt; width:14.25pt; height:14.25pt; border:solid 0.75pt #000000;\"")->get_Success());
    }
    //ExEnd
}

namespace gtest_test
{

using ExportFormFields_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::ExportFormFields)>::type;

struct ExportFormFieldsVP : public ExHtmlFixedSaveOptions, public ApiExamples::ExHtmlFixedSaveOptions, public ::testing::WithParamInterface<ExportFormFields_Args>
{
    static std::vector<ExportFormFields_Args> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExportFormFieldsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportFormFields(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlFixedSaveOptions, ExportFormFieldsVP, ::testing::ValuesIn(ExportFormFieldsVP::TestCases()));

} // namespace gtest_test

void ExHtmlFixedSaveOptions::AddCssClassNamesPrefix()
{
    //ExStart
    //ExFor:HtmlFixedSaveOptions.CssClassNamesPrefix
    //ExFor:HtmlFixedSaveOptions.SaveFontFaceCssSeparately
    //ExSummary:Shows how to add prefix to all class names in css file.
    auto doc = MakeObject<Document>(MyDir + u"Bookmarks.docx");

    auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
    htmlFixedSaveOptions->set_CssClassNamesPrefix(u"myprefix");
    htmlFixedSaveOptions->set_SaveFontFaceCssSeparately(true);

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.AddCssClassNamesPrefix.html", htmlFixedSaveOptions);

    String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.AddCssClassNamesPrefix.html");

    ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, String(u"<div class=\"myprefixdiv myprefixpage\" style=\"width:595[.]3pt; height:841[.]9pt;\">") + u"<div class=\"myprefixdiv\" style=\"left:85[.]05pt; top:36pt; clip:rect[(]0pt,510[.]25pt,74[.]95pt,-85.05pt[)];\">" + u"<span class=\"myprefixspan myprefixtext001\" style=\"font-size:11pt; left:294[.]73pt; top:0[.]36pt;\">")->get_Success());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHtmlFixedSaveOptions, AddCssClassNamesPrefix)
{
    s_instance->AddCssClassNamesPrefix();
}

} // namespace gtest_test

void ExHtmlFixedSaveOptions::HorizontalAlignment(Aspose::Words::Saving::HtmlFixedPageHorizontalAlignment pageHorizontalAlignment)
{
    //ExStart
    //ExFor:HtmlFixedSaveOptions.PageHorizontalAlignment
    //ExFor:HtmlFixedPageHorizontalAlignment
    //ExSummary:Shows how to set the horizontal alignment of pages in HTML file.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
    htmlFixedSaveOptions->set_PageHorizontalAlignment(pageHorizontalAlignment);

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.HorizontalAlignment.html", htmlFixedSaveOptions);

    String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.HorizontalAlignment/styles.css");

    switch (pageHorizontalAlignment)
    {
        case Aspose::Words::Saving::HtmlFixedPageHorizontalAlignment::Center:
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"[.]awpage { position:relative; border:solid 1pt black; margin:10pt auto 10pt auto; overflow:hidden; }")->get_Success());
            break;

        case Aspose::Words::Saving::HtmlFixedPageHorizontalAlignment::Left:
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"[.]awpage { position:relative; border:solid 1pt black; margin:10pt auto 10pt 10pt; overflow:hidden; }")->get_Success());
            break;

        case Aspose::Words::Saving::HtmlFixedPageHorizontalAlignment::Right:
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"[.]awpage { position:relative; border:solid 1pt black; margin:10pt 10pt 10pt auto; overflow:hidden; }")->get_Success());
            break;

    }
    //ExEnd
}

namespace gtest_test
{

using HorizontalAlignment_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlFixedSaveOptions::HorizontalAlignment)>::type;

struct HorizontalAlignmentVP : public ExHtmlFixedSaveOptions, public ApiExamples::ExHtmlFixedSaveOptions, public ::testing::WithParamInterface<HorizontalAlignment_Args>
{
    static std::vector<HorizontalAlignment_Args> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::HtmlFixedPageHorizontalAlignment::Center),
            std::make_tuple(Aspose::Words::Saving::HtmlFixedPageHorizontalAlignment::Left),
            std::make_tuple(Aspose::Words::Saving::HtmlFixedPageHorizontalAlignment::Right),
        };
    }
};

TEST_P(HorizontalAlignmentVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->HorizontalAlignment(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlFixedSaveOptions, HorizontalAlignmentVP, ::testing::ValuesIn(HorizontalAlignmentVP::TestCases()));

} // namespace gtest_test

void ExHtmlFixedSaveOptions::PageMargins()
{
    //ExStart
    //ExFor:HtmlFixedSaveOptions.PageMargins
    //ExSummary:Shows how to set the margins around pages in HTML file.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    auto saveOptions = MakeObject<HtmlFixedSaveOptions>();
    saveOptions->set_PageMargins(15);

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.PageMargins.html", saveOptions);

    String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.PageMargins/styles.css");

    ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"[.]awpage { position:relative; border:solid 1pt black; margin:15pt auto 15pt auto; overflow:hidden; }")->get_Success());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHtmlFixedSaveOptions, PageMargins)
{
    s_instance->PageMargins();
}

} // namespace gtest_test

void ExHtmlFixedSaveOptions::PageMarginsException()
{
    auto saveOptions = MakeObject<HtmlFixedSaveOptions>();

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_2 = [&saveOptions]()
    {
        saveOptions->set_PageMargins(-1);
    };

    ASSERT_THROW(_local_func_2(), System::ArgumentException);
}

namespace gtest_test
{

TEST_F(ExHtmlFixedSaveOptions, PageMarginsException)
{
    s_instance->PageMarginsException();
}

} // namespace gtest_test

void ExHtmlFixedSaveOptions::OptimizeGraphicsOutput()
{
    //ExStart
    //ExFor:FixedPageSaveOptions.OptimizeOutput
    //ExFor:HtmlFixedSaveOptions.OptimizeOutput
    //ExSummary:Shows how to optimize document objects while saving to html.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    auto saveOptions = MakeObject<HtmlFixedSaveOptions>();
    saveOptions->set_OptimizeOutput(false);

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.OptimizeGraphicsOutput.Unoptimized.html", saveOptions);

    saveOptions->set_OptimizeOutput(true);

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.OptimizeGraphicsOutput.Optimized.html", saveOptions);

    ASSERT_TRUE(MakeObject<System::IO::FileInfo>(ArtifactsDir + u"HtmlFixedSaveOptions.OptimizeGraphicsOutput.Unoptimized.html")->get_Length() > MakeObject<System::IO::FileInfo>(ArtifactsDir + u"HtmlFixedSaveOptions.OptimizeGraphicsOutput.Optimized.html")->get_Length());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHtmlFixedSaveOptions, OptimizeGraphicsOutput)
{
    s_instance->OptimizeGraphicsOutput();
}

} // namespace gtest_test

void ExHtmlFixedSaveOptions::UsingMachineFonts()
{
    auto doc = MakeObject<Document>(MyDir + u"Bullet points with alternative font.docx");

    auto saveOptions = MakeObject<HtmlFixedSaveOptions>();
    saveOptions->set_ExportEmbeddedCss(true);
    saveOptions->set_UseTargetMachineFonts(true);
    saveOptions->set_FontFormat(Aspose::Words::Saving::ExportFontFormat::Ttf);
    saveOptions->set_ExportEmbeddedFonts(false);
    saveOptions->set_ResourceSavingCallback(MakeObject<ExHtmlFixedSaveOptions::ResourceSavingCallback>());

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);

    String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlFixedSaveOptions.UsingMachineFonts.html");

    if (saveOptions->get_UseTargetMachineFonts())
    {
        ASSERT_FALSE(System::Text::RegularExpressions::Regex::Match(outDocContents, u"@font-face")->get_Success());
    }
    else
    {
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, String(u"@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], ") + u"url[(]'HtmlFixedSaveOptions.UsingMachineFonts/font001.ttf'[)] format[(]'truetype'[)]; }")->get_Success());
    }
}

namespace gtest_test
{

TEST_F(ExHtmlFixedSaveOptions, UsingMachineFonts)
{
    s_instance->UsingMachineFonts();
}

} // namespace gtest_test

void ExHtmlFixedSaveOptions::HtmlFixedResourceFolder()
{
    // Open a document which contains images
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    auto options = MakeObject<HtmlFixedSaveOptions>();
    options->set_SaveFormat(Aspose::Words::SaveFormat::HtmlFixed);
    options->set_ExportEmbeddedImages(false);
    options->set_ResourcesFolder(ArtifactsDir + u"HtmlFixedResourceFolder");
    options->set_ResourcesFolderAlias(ArtifactsDir + u"HtmlFixedResourceFolderAlias");
    options->set_ShowPageBorder(false);
    options->set_ResourceSavingCallback(MakeObject<ExHtmlFixedSaveOptions::ResourceUriPrinter>());

    // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder
    // We must ensure the folder exists before the streams can put their resources into it
    System::IO::Directory::CreateDirectory_(options->get_ResourcesFolderAlias());

    doc->Save(ArtifactsDir + u"HtmlFixedSaveOptions.HtmlFixedResourceFolder.html", options);

    ArrayPtr<String> resourceFiles = System::IO::Directory::GetFiles(ArtifactsDir + u"HtmlFixedResourceFolderAlias");

    ASSERT_FALSE(System::IO::Directory::Exists(ArtifactsDir + u"HtmlFixedResourceFolder"));

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(String f)> _local_func_3 = [](String f)
    {
        return f.EndsWith(u".jpeg") || f.EndsWith(u".png") || f.EndsWith(u".css");
    };

    ASSERT_EQ(6, resourceFiles->LINQ_Count(static_cast<System::Func<String, bool>>(_local_func_3)));
}

namespace gtest_test
{

TEST_F(ExHtmlFixedSaveOptions, HtmlFixedResourceFolder)
{
    s_instance->HtmlFixedResourceFolder();
}

} // namespace gtest_test

} // namespace ApiExamples
