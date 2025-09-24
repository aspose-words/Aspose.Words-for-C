// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHtmlSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/timespan.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/primitive_types.h>
#include <system/object_ext.h>
#include <system/nullable.h>
#include <system/math.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/search_option.h>
#include <system/io/path.h>
#include <system/io/memory_stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file_info.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/globalization/culture_info.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/environment.h>
#include <system/enum_helpers.h>
#include <system/default.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/rectangle_f.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PaperSize.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlElementSizeOutputMode.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentSplitCriteria.h>
#include <Aspose.Words.Cpp/Model/Saving/CssStyleSheetType.h>
#include <Aspose.Words.Cpp/Model/Loading/HtmlLoadOptions.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSubstitutionSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontInfo.h>
#include <Aspose.Words.Cpp/Model/Fonts/DefaultFontSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldToc.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BuildVersionInfo.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutCollector.h>

#include "TestUtil.h"
#include "DocumentHelper.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(557788919u, ::Aspose::Words::ApiExamples::ExHtmlSaveOptions::HandleFontSaving, ThisTypeBaseTypesInfo);

void ExHtmlSaveOptions::HandleFontSaving::FontSaving(System::SharedPtr<Aspose::Words::Saving::FontSavingArgs> args)
{
    std::cout << System::String::Format(u"Font:\t{0}", args->get_FontFamilyName());
    if (args->get_Bold())
    {
        std::cout << ", bold";
    }
    if (args->get_Italic())
    {
        std::cout << ", italic";
    }
    std::cout << System::String::Format(u"\nSource:\t{0}, {1} bytes\n", args->get_OriginalFileName(), args->get_OriginalFileSize()) << std::endl;
    
    // We can also access the source document from here.
    ASSERT_TRUE(args->get_Document()->get_OriginalFileName().EndsWith(u"Rendering.docx"));
    
    ASSERT_TRUE(args->get_IsExportNeeded());
    ASSERT_TRUE(args->get_IsSubsettingNeeded());
    
    // There are two ways of saving an exported font.
    // 1 -  Save it to a local file system location:
    args->set_FontFileName(args->get_OriginalFileName().Split(System::IO::Path::DirectorySeparatorChar)->LINQ_Last());
    
    // 2 -  Save it to a stream:
    args->set_FontStream(System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + args->get_OriginalFileName().Split(System::IO::Path::DirectorySeparatorChar)->LINQ_Last(), System::IO::FileMode::Create));
    ASSERT_FALSE(args->get_KeepFontStreamOpen());
}

RTTI_INFO_IMPL_HASH(2128611150u, ::Aspose::Words::ApiExamples::ExHtmlSaveOptions::ImageShapePrinter, ThisTypeBaseTypesInfo);

void ExHtmlSaveOptions::ImageShapePrinter::ImageSaving(System::SharedPtr<Aspose::Words::Saving::ImageSavingArgs> args)
{
    args->set_KeepImageStreamOpen(false);
    ASSERT_TRUE(args->get_IsImageAvailable());
    
    std::cout << System::String::Format(u"{0} Image #{1}", args->get_Document()->get_OriginalFileName().Split(u'\\')->LINQ_Last(), ++mImageCount) << std::endl;
    
    auto layoutCollector = System::MakeObject<Aspose::Words::Layout::LayoutCollector>(args->get_Document());
    
    std::cout << System::String::Format(u"\tOn page:\t{0}", layoutCollector->GetStartPageIndex(args->get_CurrentShape())) << std::endl;
    std::cout << System::String::Format(u"\tDimensions:\t{0}", args->get_CurrentShape()->get_Bounds()) << std::endl;
    std::cout << System::String::Format(u"\tAlignment:\t{0}", args->get_CurrentShape()->get_VerticalAlignment()) << std::endl;
    std::cout << System::String::Format(u"\tWrap type:\t{0}", args->get_CurrentShape()->get_WrapType()) << std::endl;
    std::cout << System::String::Format(u"Output filename:\t{0}\n", args->get_ImageFileName()) << std::endl;
}

ExHtmlSaveOptions::ImageShapePrinter::ImageShapePrinter() : mImageCount(0)
{
}

RTTI_INFO_IMPL_HASH(116942450u, ::Aspose::Words::ApiExamples::ExHtmlSaveOptions::SavingProgressCallback, ThisTypeBaseTypesInfo);

const double ExHtmlSaveOptions::SavingProgressCallback::MaxDuration = 0.1;

ExHtmlSaveOptions::SavingProgressCallback::SavingProgressCallback()
{
    mSavingStartedAt = System::DateTime::get_Now();
}

void ExHtmlSaveOptions::SavingProgressCallback::Notify(System::SharedPtr<Aspose::Words::Saving::DocumentSavingArgs> args)
{
    System::DateTime canceledAt = System::DateTime::get_Now();
    double ellapsedSeconds = (canceledAt - mSavingStartedAt).get_TotalSeconds();
    if (ellapsedSeconds > MaxDuration)
    {
        throw System::OperationCanceledException(System::String::Format(u"EstimatedProgress = {0}; CanceledAt = {1}", args->get_EstimatedProgress(), canceledAt));
    }
}


RTTI_INFO_IMPL_HASH(1344508550u, ::Aspose::Words::ApiExamples::ExHtmlSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExHtmlSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExHtmlSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExHtmlSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExHtmlSaveOptions> ExHtmlSaveOptions::s_instance;

} // namespace gtest_test

void ExHtmlSaveOptions::ExportPageMarginsEpub(Aspose::Words::SaveFormat saveFormat)
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"TextBoxes.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_SaveFormat(saveFormat);
    saveOptions->set_ExportPageMargins(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportPageMarginsEpub" + Aspose::Words::FileFormatUtil::SaveFormatToExtension(saveFormat), saveOptions);
}

namespace gtest_test
{

using ExHtmlSaveOptions_ExportPageMarginsEpub_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ExportPageMarginsEpub)>::type;

struct ExHtmlSaveOptions_ExportPageMarginsEpub : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportPageMarginsEpub_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::SaveFormat::Html),
            std::make_tuple(Aspose::Words::SaveFormat::Mhtml),
            std::make_tuple(Aspose::Words::SaveFormat::Epub),
            std::make_tuple(Aspose::Words::SaveFormat::Azw3),
            std::make_tuple(Aspose::Words::SaveFormat::Mobi),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportPageMarginsEpub, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportPageMarginsEpub(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportPageMarginsEpub, ::testing::ValuesIn(ExHtmlSaveOptions_ExportPageMarginsEpub::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ExportOfficeMathEpub(Aspose::Words::SaveFormat saveFormat, Aspose::Words::Saving::HtmlOfficeMathOutputMode outputMode)
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_OfficeMathOutputMode(outputMode);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportOfficeMathEpub" + Aspose::Words::FileFormatUtil::SaveFormatToExtension(saveFormat), saveOptions);
}

namespace gtest_test
{

using ExHtmlSaveOptions_ExportOfficeMathEpub_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ExportOfficeMathEpub)>::type;

struct ExHtmlSaveOptions_ExportOfficeMathEpub : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportOfficeMathEpub_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::SaveFormat::Html, Aspose::Words::Saving::HtmlOfficeMathOutputMode::Image),
            std::make_tuple(Aspose::Words::SaveFormat::Mhtml, Aspose::Words::Saving::HtmlOfficeMathOutputMode::MathML),
            std::make_tuple(Aspose::Words::SaveFormat::Epub, Aspose::Words::Saving::HtmlOfficeMathOutputMode::Text),
            std::make_tuple(Aspose::Words::SaveFormat::Azw3, Aspose::Words::Saving::HtmlOfficeMathOutputMode::Text),
            std::make_tuple(Aspose::Words::SaveFormat::Mobi, Aspose::Words::Saving::HtmlOfficeMathOutputMode::Text),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportOfficeMathEpub, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportOfficeMathEpub(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportOfficeMathEpub, ::testing::ValuesIn(ExHtmlSaveOptions_ExportOfficeMathEpub::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ExportTextBoxAsSvgEpub(Aspose::Words::SaveFormat saveFormat, bool isTextBoxAsSvg)
{
    System::ArrayPtr<System::String> dirFiles;
    
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> textbox = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 300, 100);
    builder->MoveTo(textbox->get_FirstParagraph());
    builder->Write(u"Hello world!");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(saveFormat);
    saveOptions->set_ExportShapesAsSvg(isTextBoxAsSvg);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportTextBoxAsSvgEpub" + Aspose::Words::FileFormatUtil::SaveFormatToExtension(saveFormat), saveOptions);
    
    switch (saveFormat)
    {
        case Aspose::Words::SaveFormat::Html:
            dirFiles = System::IO::Directory::GetFiles(get_ArtifactsDir(), u"HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png", System::IO::SearchOption::AllDirectories);
            ASSERT_EQ(0, dirFiles->get_Length());
            return;
        
        case Aspose::Words::SaveFormat::Epub:
            dirFiles = System::IO::Directory::GetFiles(get_ArtifactsDir(), u"HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png", System::IO::SearchOption::AllDirectories);
            ASSERT_EQ(0, dirFiles->get_Length());
            return;
        
        case Aspose::Words::SaveFormat::Mhtml:
            dirFiles = System::IO::Directory::GetFiles(get_ArtifactsDir(), u"HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png", System::IO::SearchOption::AllDirectories);
            ASSERT_EQ(0, dirFiles->get_Length());
            return;
        
        case Aspose::Words::SaveFormat::Azw3:
            dirFiles = System::IO::Directory::GetFiles(get_ArtifactsDir(), u"HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png", System::IO::SearchOption::AllDirectories);
            ASSERT_EQ(0, dirFiles->get_Length());
            return;
        
        case Aspose::Words::SaveFormat::Mobi:
            dirFiles = System::IO::Directory::GetFiles(get_ArtifactsDir(), u"HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png", System::IO::SearchOption::AllDirectories);
            ASSERT_EQ(0, dirFiles->get_Length());
            return;
        
        default:
            break;
    }
}

namespace gtest_test
{

using ExHtmlSaveOptions_ExportTextBoxAsSvgEpub_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ExportTextBoxAsSvgEpub)>::type;

struct ExHtmlSaveOptions_ExportTextBoxAsSvgEpub : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportTextBoxAsSvgEpub_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::SaveFormat::Html, true),
            std::make_tuple(Aspose::Words::SaveFormat::Epub, true),
            std::make_tuple(Aspose::Words::SaveFormat::Mhtml, false),
            std::make_tuple(Aspose::Words::SaveFormat::Azw3, false),
            std::make_tuple(Aspose::Words::SaveFormat::Mobi, false),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ExportTextBoxAsSvgEpub, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportTextBoxAsSvgEpub(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportTextBoxAsSvgEpub, ::testing::ValuesIn(ExHtmlSaveOptions_ExportTextBoxAsSvgEpub::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::CreateAZW3Toc()
{
    //ExStart
    //ExFor:HtmlSaveOptions.NavigationMapLevel
    //ExSummary:Shows how to generate table of contents for Azw3 documents.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Azw3);
    options->set_NavigationMapLevel(2);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.CreateAZW3Toc.azw3", options);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, CreateAZW3Toc)
{
    s_instance->CreateAZW3Toc();
}

} // namespace gtest_test

void ExHtmlSaveOptions::CreateMobiToc()
{
    //ExStart
    //ExFor:HtmlSaveOptions.NavigationMapLevel
    //ExSummary:Shows how to generate table of contents for Mobi documents.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Mobi);
    options->set_NavigationMapLevel(5);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.CreateMobiToc.mobi", options);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, CreateMobiToc)
{
    s_instance->CreateMobiToc();
}

} // namespace gtest_test

void ExHtmlSaveOptions::ControlListLabelsExport(Aspose::Words::Saving::ExportListLabels howExportListLabels)
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Lists::List> bulletedList = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::BulletDefault);
    builder->get_ListFormat()->set_List(bulletedList);
    builder->get_ParagraphFormat()->set_LeftIndent(72);
    builder->Writeln(u"Bulleted list item 1.");
    builder->Writeln(u"Bulleted list item 2.");
    builder->get_ParagraphFormat()->ClearFormatting();
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Html);
    // 'ExportListLabels.Auto' - this option uses <ul> and <ol> tags are used for list label representation if it does not cause formatting loss, 
    // otherwise HTML <p> tag is used. This is also the default value.
    // 'ExportListLabels.AsInlineText' - using this option the <p> tag is used for any list label representation.
    // 'ExportListLabels.ByHtmlTags' - The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
    saveOptions->set_ExportListLabels(howExportListLabels);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ControlListLabelsExport.html", saveOptions);
}

namespace gtest_test
{

using ExHtmlSaveOptions_ControlListLabelsExport_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ControlListLabelsExport)>::type;

struct ExHtmlSaveOptions_ControlListLabelsExport : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ControlListLabelsExport_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::ExportListLabels::Auto),
            std::make_tuple(Aspose::Words::Saving::ExportListLabels::AsInlineText),
            std::make_tuple(Aspose::Words::Saving::ExportListLabels::ByHtmlTags),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ControlListLabelsExport, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ControlListLabelsExport(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ControlListLabelsExport, ::testing::ValuesIn(ExHtmlSaveOptions_ControlListLabelsExport::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ExportUrlForLinkedImage(bool export_)
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Linked image.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_ExportOriginalUrlForLinkedImages(export_);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);
    
    System::ArrayPtr<System::String> dirFiles = System::IO::Directory::GetFiles(get_ArtifactsDir(), u"HtmlSaveOptions.ExportUrlForLinkedImage.001.png", System::IO::SearchOption::AllDirectories);
    
    Aspose::Words::ApiExamples::DocumentHelper::FindTextInFile(get_ArtifactsDir() + u"HtmlSaveOptions.ExportUrlForLinkedImage.html", dirFiles->get_Length() == 0 ? System::String(u"<img src=\"http://www.aspose.com/images/aspose-logo.gif\"") : System::String(u"<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\""));
}

namespace gtest_test
{

using ExHtmlSaveOptions_ExportUrlForLinkedImage_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ExportUrlForLinkedImage)>::type;

struct ExHtmlSaveOptions_ExportUrlForLinkedImage : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportUrlForLinkedImage_Args>
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

TEST_P(ExHtmlSaveOptions_ExportUrlForLinkedImage, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportUrlForLinkedImage(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportUrlForLinkedImage, ::testing::ValuesIn(ExHtmlSaveOptions_ExportUrlForLinkedImage::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ExportRoundtripInformation()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"TextBoxes.docx");
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_ExportRoundtripInformation(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.RoundtripInformation.html", saveOptions);
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, ExportRoundtripInformation)
{
    s_instance->ExportRoundtripInformation();
}

} // namespace gtest_test

void ExHtmlSaveOptions::RoundtripInformationDefaulValue()
{
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Html);
    ASPOSE_ASSERT_EQ(true, saveOptions->get_ExportRoundtripInformation());
    
    saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Mhtml);
    ASPOSE_ASSERT_EQ(false, saveOptions->get_ExportRoundtripInformation());
    
    saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Epub);
    ASPOSE_ASSERT_EQ(false, saveOptions->get_ExportRoundtripInformation());
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, RoundtripInformationDefaulValue)
{
    s_instance->RoundtripInformationDefaulValue();
}

} // namespace gtest_test

void ExHtmlSaveOptions::ExternalResourceSavingConfig()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_CssStyleSheetType(Aspose::Words::Saving::CssStyleSheetType::External);
    saveOptions->set_ExportFontResources(true);
    saveOptions->set_ResourceFolder(u"Resources");
    saveOptions->set_ResourceFolderAlias(u"https://www.aspose.com/");
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExternalResourceSavingConfig.html", saveOptions);
    
    System::ArrayPtr<System::String> imageFiles = System::IO::Directory::GetFiles(get_ArtifactsDir() + u"Resources/", u"HtmlSaveOptions.ExternalResourceSavingConfig*.png", System::IO::SearchOption::AllDirectories);
    ASSERT_EQ(8, imageFiles->get_Length());
    
    System::ArrayPtr<System::String> fontFiles = System::IO::Directory::GetFiles(get_ArtifactsDir() + u"Resources/", u"HtmlSaveOptions.ExternalResourceSavingConfig*.ttf", System::IO::SearchOption::AllDirectories);
    ASSERT_EQ(10, fontFiles->get_Length());
    
    System::ArrayPtr<System::String> cssFiles = System::IO::Directory::GetFiles(get_ArtifactsDir() + u"Resources/", u"HtmlSaveOptions.ExternalResourceSavingConfig*.css", System::IO::SearchOption::AllDirectories);
    ASSERT_EQ(1, cssFiles->get_Length());
    
    Aspose::Words::ApiExamples::DocumentHelper::FindTextInFile(get_ArtifactsDir() + u"HtmlSaveOptions.ExternalResourceSavingConfig.html", u"<link href=\"https://www.aspose.com/HtmlSaveOptions.ExternalResourceSavingConfig.css\"");
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, ExternalResourceSavingConfig)
{
    s_instance->ExternalResourceSavingConfig();
}

} // namespace gtest_test

void ExHtmlSaveOptions::ConvertFontsAsBase64()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"TextBoxes.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_CssStyleSheetType(Aspose::Words::Saving::CssStyleSheetType::External);
    saveOptions->set_ResourceFolder(u"Resources");
    saveOptions->set_ExportFontResources(true);
    saveOptions->set_ExportFontsAsBase64(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ConvertFontsAsBase64.html", saveOptions);
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, ConvertFontsAsBase64)
{
    s_instance->ConvertFontsAsBase64();
}

} // namespace gtest_test

void ExHtmlSaveOptions::Html5Support(Aspose::Words::Saving::HtmlVersion htmlVersion)
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_HtmlVersion(htmlVersion);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.Html5Support.html", saveOptions);
}

namespace gtest_test
{

using ExHtmlSaveOptions_Html5Support_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::Html5Support)>::type;

struct ExHtmlSaveOptions_Html5Support : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_Html5Support_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::HtmlVersion::Html5),
            std::make_tuple(Aspose::Words::Saving::HtmlVersion::Xhtml),
        };
    }
};

TEST_P(ExHtmlSaveOptions_Html5Support, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Html5Support(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_Html5Support, ::testing::ValuesIn(ExHtmlSaveOptions_Html5Support::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ExportFonts(bool exportAsBase64)
{
    System::String fontsFolder = get_ArtifactsDir() + u"HtmlSaveOptions.ExportFonts.Resources";
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_ExportFontResources(true);
    saveOptions->set_FontsFolder(fontsFolder);
    saveOptions->set_ExportFontsAsBase64(exportAsBase64);
    
    switch (exportAsBase64)
    {
        case false:
            doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportFonts.False.html", saveOptions);
            ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(fontsFolder, u"HtmlSaveOptions.ExportFonts.False.times.ttf", System::IO::SearchOption::AllDirectories)));
            System::IO::Directory::Delete(fontsFolder, true);
            break;
        
        case true:
            doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportFonts.True.html", saveOptions);
            ASSERT_FALSE(System::IO::Directory::Exists(fontsFolder));
            break;
        
    }
}

namespace gtest_test
{

using ExHtmlSaveOptions_ExportFonts_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ExportFonts)>::type;

struct ExHtmlSaveOptions_ExportFonts : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportFonts_Args>
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

TEST_P(ExHtmlSaveOptions_ExportFonts, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportFonts(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportFonts, ::testing::ValuesIn(ExHtmlSaveOptions_ExportFonts::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ResourceFolderPriority()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_CssStyleSheetType(Aspose::Words::Saving::CssStyleSheetType::External);
    saveOptions->set_ExportFontResources(true);
    saveOptions->set_ResourceFolder(get_ArtifactsDir() + u"Resources");
    saveOptions->set_ResourceFolderAlias(u"http://example.com/resources");
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ResourceFolderPriority.html", saveOptions);
    
    ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(get_ArtifactsDir() + u"Resources", u"HtmlSaveOptions.ResourceFolderPriority.001.png", System::IO::SearchOption::AllDirectories)));
    ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(get_ArtifactsDir() + u"Resources", u"HtmlSaveOptions.ResourceFolderPriority.002.png", System::IO::SearchOption::AllDirectories)));
    ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(get_ArtifactsDir() + u"Resources", u"HtmlSaveOptions.ResourceFolderPriority.arial.ttf", System::IO::SearchOption::AllDirectories)));
    ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(get_ArtifactsDir() + u"Resources", u"HtmlSaveOptions.ResourceFolderPriority.css", System::IO::SearchOption::AllDirectories)));
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, ResourceFolderPriority)
{
    s_instance->ResourceFolderPriority();
}

} // namespace gtest_test

void ExHtmlSaveOptions::ResourceFolderLowPriority()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_CssStyleSheetType(Aspose::Words::Saving::CssStyleSheetType::External);
    saveOptions->set_ExportFontResources(true);
    saveOptions->set_FontsFolder(get_ArtifactsDir() + u"Fonts");
    saveOptions->set_ImagesFolder(get_ArtifactsDir() + u"Images");
    saveOptions->set_ResourceFolder(get_ArtifactsDir() + u"Resources");
    saveOptions->set_ResourceFolderAlias(u"http://example.com/resources");
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ResourceFolderLowPriority.html", saveOptions);
    
    ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(get_ArtifactsDir() + u"Images", u"HtmlSaveOptions.ResourceFolderLowPriority.001.png", System::IO::SearchOption::AllDirectories)));
    ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(get_ArtifactsDir() + u"Images", u"HtmlSaveOptions.ResourceFolderLowPriority.002.png", System::IO::SearchOption::AllDirectories)));
    ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(get_ArtifactsDir() + u"Fonts", u"HtmlSaveOptions.ResourceFolderLowPriority.arial.ttf", System::IO::SearchOption::AllDirectories)));
    ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(get_ArtifactsDir() + u"Resources", u"HtmlSaveOptions.ResourceFolderLowPriority.css", System::IO::SearchOption::AllDirectories)));
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, ResourceFolderLowPriority)
{
    s_instance->ResourceFolderLowPriority();
}

} // namespace gtest_test

void ExHtmlSaveOptions::SvgMetafileFormat()
{
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    builder->Write(u"Here is an SVG image: ");
    builder->InsertHtml(u"<svg height='210' width='500'>\r\n                    <polygon points='100,10 40,198 190,78 10,78 160,198' \r\n                        style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n                  </svg> ");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_MetafileFormat(Aspose::Words::Saving::HtmlMetafileFormat::Png);
    builder->get_Document()->Save(get_ArtifactsDir() + u"HtmlSaveOptions.SvgMetafileFormat.html", saveOptions);
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, SvgMetafileFormat)
{
    s_instance->SvgMetafileFormat();
}

} // namespace gtest_test

void ExHtmlSaveOptions::PngMetafileFormat()
{
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    builder->Write(u"Here is an Png image: ");
    builder->InsertHtml(u"<svg height='210' width='500'>\r\n                    <polygon points='100,10 40,198 190,78 10,78 160,198' \r\n                        style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n                  </svg> ");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_MetafileFormat(Aspose::Words::Saving::HtmlMetafileFormat::Png);
    builder->get_Document()->Save(get_ArtifactsDir() + u"HtmlSaveOptions.PngMetafileFormat.html", saveOptions);
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, PngMetafileFormat)
{
    s_instance->PngMetafileFormat();
}

} // namespace gtest_test

void ExHtmlSaveOptions::EmfOrWmfMetafileFormat()
{
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    builder->Write(u"Here is an image as is: ");
    builder->InsertHtml(u"<img src=\"data:image/png;base64,\r\n                    iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP\r\n                    C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA\r\n                    AAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J\r\n                    REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq\r\n                    ch9//q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0\r\n                    vr4MkhoXe0rZigAAAABJRU5ErkJggg==\" alt=\"Red dot\" />");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_MetafileFormat(Aspose::Words::Saving::HtmlMetafileFormat::EmfOrWmf);
    builder->get_Document()->Save(get_ArtifactsDir() + u"HtmlSaveOptions.EmfOrWmfMetafileFormat.html", saveOptions);
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, EmfOrWmfMetafileFormat)
{
    s_instance->EmfOrWmfMetafileFormat();
}

} // namespace gtest_test

void ExHtmlSaveOptions::CssClassNamesPrefix()
{
    //ExStart
    //ExFor:HtmlSaveOptions.CssClassNamePrefix
    //ExSummary:Shows how to save a document to HTML, and add a prefix to all of its CSS class names.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Paragraphs.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_CssStyleSheetType(Aspose::Words::Saving::CssStyleSheetType::External);
    saveOptions->set_CssClassNamePrefix(u"myprefix-");
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.CssClassNamePrefix.html");
    
    ASSERT_TRUE(outDocContents.Contains(u"<p class=\"myprefix-Header\">"));
    ASSERT_TRUE(outDocContents.Contains(u"<p class=\"myprefix-Footer\">"));
    
    outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.CssClassNamePrefix.css");
    
    ASSERT_TRUE(outDocContents.Contains(u".myprefix-Footer { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt; -aw-style-name:footer }"));
    ASSERT_TRUE(outDocContents.Contains(u".myprefix-Header { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt; -aw-style-name:header }"));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, CssClassNamesPrefix)
{
    s_instance->CssClassNamesPrefix();
}

} // namespace gtest_test

void ExHtmlSaveOptions::CssClassNamesNotValidPrefix()
{
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    ASSERT_THROW(static_cast<std::function<void()>>([&saveOptions]() -> void
    {
        saveOptions->set_CssClassNamePrefix(u"@%-");
    })(), System::ArgumentException) << "The class name prefix must be a valid CSS identifier.";
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, CssClassNamesNotValidPrefix)
{
    s_instance->CssClassNamesNotValidPrefix();
}

} // namespace gtest_test

void ExHtmlSaveOptions::CssClassNamesNullPrefix()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Paragraphs.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_CssStyleSheetType(Aspose::Words::Saving::CssStyleSheetType::Embedded);
    saveOptions->set_CssClassNamePrefix(nullptr);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, CssClassNamesNullPrefix)
{
    s_instance->CssClassNamesNullPrefix();
}

} // namespace gtest_test

void ExHtmlSaveOptions::ContentIdScheme()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Mhtml);
    saveOptions->set_PrettyFormat(true);
    saveOptions->set_ExportCidUrlsForMhtmlResources(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ContentIdScheme.mhtml", saveOptions);
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, ContentIdScheme)
{
    s_instance->ContentIdScheme();
}

} // namespace gtest_test

void ExHtmlSaveOptions::ResolveFontNames(bool resolveFontNames)
{
    //ExStart
    //ExFor:HtmlSaveOptions.ResolveFontNames
    //ExSummary:Shows how to resolve all font names before writing them to HTML.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Missing font.docx");
    
    // This document contains text that names a font that we do not have.
    ASSERT_FALSE(System::TestTools::IsNull(doc->get_FontInfos()->idx_get(u"28 Days Later")));
    
    // If we have no way of getting this font, and we want to be able to display all the text
    // in this document in an output HTML, we can substitute it with another font.
    auto fontSettings = System::MakeObject<Aspose::Words::Fonts::FontSettings>();
    fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
    fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_Enabled(true);
    
    doc->set_FontSettings(fontSettings);
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Html);
    // By default, this option is set to 'False' and Aspose.Words writes font names as specified in the source document
    saveOptions->set_ResolveFontNames(resolveFontNames);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ResolveFontNames.html", saveOptions);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.ResolveFontNames.html");
    
    ASSERT_TRUE(resolveFontNames ? System::Text::RegularExpressions::Regex::Match(outDocContents, u"<span style=\"font-family:Arial\">")->get_Success() : System::Text::RegularExpressions::Regex::Match(outDocContents, u"<span style=\"font-family:\'28 Days Later\'\">")->get_Success());
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_ResolveFontNames_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ResolveFontNames)>::type;

struct ExHtmlSaveOptions_ResolveFontNames : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ResolveFontNames_Args>
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

TEST_P(ExHtmlSaveOptions_ResolveFontNames, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ResolveFontNames(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(DISABLED_, ExHtmlSaveOptions_ResolveFontNames, ::testing::ValuesIn(ExHtmlSaveOptions_ResolveFontNames::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::HeadingLevels()
{
    //ExStart
    //ExFor:HtmlSaveOptions.DocumentSplitHeadingLevel
    //ExSummary:Shows how to split an output HTML document by headings into several parts.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Every paragraph that we format using a "Heading" style can serve as a heading.
    // Each heading may also have a heading level, determined by the number of its heading style.
    // The headings below are of levels 1-3.
    builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(u"Heading 1"));
    builder->Writeln(u"Heading #1");
    builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(u"Heading 2"));
    builder->Writeln(u"Heading #2");
    builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(u"Heading 3"));
    builder->Writeln(u"Heading #3");
    builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(u"Heading 1"));
    builder->Writeln(u"Heading #4");
    builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(u"Heading 2"));
    builder->Writeln(u"Heading #5");
    builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(u"Heading 3"));
    builder->Writeln(u"Heading #6");
    
    // Create a HtmlSaveOptions object and set the split criteria to "HeadingParagraph".
    // These criteria will split the document at paragraphs with "Heading" styles into several smaller documents,
    // and save each document in a separate HTML file in the local file system.
    // We will also set the maximum heading level, which splits the document to 2.
    // Saving the document will split it at headings of levels 1 and 2, but not at 3 to 9.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_DocumentSplitCriteria(Aspose::Words::Saving::DocumentSplitCriteria::HeadingParagraph);
    options->set_DocumentSplitHeadingLevel(2);
    
    // Our document has four headings of levels 1 - 2. One of those headings will not be
    // a split point since it is at the beginning of the document.
    // The saving operation will split our document at three places, into four smaller documents.
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.HeadingLevels.html", options);
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"HtmlSaveOptions.HeadingLevels.html");
    
    ASSERT_EQ(u"Heading #1", doc->GetText().Trim());
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"HtmlSaveOptions.HeadingLevels-01.html");
    
    ASSERT_EQ(System::String(u"Heading #2\r") + u"Heading #3", doc->GetText().Trim());
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"HtmlSaveOptions.HeadingLevels-02.html");
    
    ASSERT_EQ(u"Heading #4", doc->GetText().Trim());
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"HtmlSaveOptions.HeadingLevels-03.html");
    
    ASSERT_EQ(System::String(u"Heading #5\r") + u"Heading #6", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, HeadingLevels)
{
    s_instance->HeadingLevels();
}

} // namespace gtest_test

void ExHtmlSaveOptions::NegativeIndent(bool allowNegativeIndent)
{
    //ExStart
    //ExFor:HtmlElementSizeOutputMode
    //ExFor:HtmlSaveOptions.AllowNegativeIndent
    //ExFor:HtmlSaveOptions.TableWidthOutputMode
    //ExSummary:Shows how to preserve negative indents in the output .html.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a table with a negative indent, which will push it to the left past the left page boundary.
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Row 1, Cell 1");
    builder->InsertCell();
    builder->Write(u"Row 1, Cell 2");
    builder->EndTable();
    table->set_LeftIndent(-36);
    table->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPoints(144));
    
    builder->InsertBreak(Aspose::Words::BreakType::ParagraphBreak);
    
    // Insert a table with a positive indent, which will push the table to the right.
    table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Row 1, Cell 1");
    builder->InsertCell();
    builder->Write(u"Row 1, Cell 2");
    builder->EndTable();
    table->set_LeftIndent(36);
    table->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPoints(144));
    
    // When we save a document to HTML, Aspose.Words will only preserve negative indents
    // such as the one we have applied to the first table if we set the "AllowNegativeIndent" flag
    // in a SaveOptions object that we will pass to "true".
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Html);
    options->set_AllowNegativeIndent(allowNegativeIndent);
    options->set_TableWidthOutputMode(Aspose::Words::Saving::HtmlElementSizeOutputMode::RelativeOnly);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.NegativeIndent.html", options);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.NegativeIndent.html");
    
    if (allowNegativeIndent)
    {
        ASSERT_TRUE(outDocContents.Contains(u"<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:-41.65pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
        ASSERT_TRUE(outDocContents.Contains(u"<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
    }
    else
    {
        ASSERT_TRUE(outDocContents.Contains(u"<table cellspacing=\"0\" cellpadding=\"0\" style=\"border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
        ASSERT_TRUE(outDocContents.Contains(u"<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_NegativeIndent_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::NegativeIndent)>::type;

struct ExHtmlSaveOptions_NegativeIndent : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_NegativeIndent_Args>
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

TEST_P(ExHtmlSaveOptions_NegativeIndent, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->NegativeIndent(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_NegativeIndent, ::testing::ValuesIn(ExHtmlSaveOptions_NegativeIndent::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::FolderAlias()
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportOriginalUrlForLinkedImages
    //ExFor:HtmlSaveOptions.FontsFolder
    //ExFor:HtmlSaveOptions.FontsFolderAlias
    //ExFor:HtmlSaveOptions.ImageResolution
    //ExFor:HtmlSaveOptions.ImagesFolderAlias
    //ExFor:HtmlSaveOptions.ResourceFolder
    //ExFor:HtmlSaveOptions.ResourceFolderAlias
    //ExSummary:Shows how to set folders and folder aliases for externally saved resources that Aspose.Words will create when saving a document to HTML.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_CssStyleSheetType(Aspose::Words::Saving::CssStyleSheetType::External);
    options->set_ExportFontResources(true);
    options->set_ImageResolution(72);
    options->set_FontResourcesSubsettingSizeThreshold(0);
    options->set_FontsFolder(get_ArtifactsDir() + u"Fonts");
    options->set_ImagesFolder(get_ArtifactsDir() + u"Images");
    options->set_ResourceFolder(get_ArtifactsDir() + u"Resources");
    options->set_FontsFolderAlias(u"http://example.com/fonts");
    options->set_ImagesFolderAlias(u"http://example.com/images");
    options->set_ResourceFolderAlias(u"http://example.com/resources");
    options->set_ExportOriginalUrlForLinkedImages(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.FolderAlias.html", options);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, FolderAlias)
{
    s_instance->FolderAlias();
}

} // namespace gtest_test

void ExHtmlSaveOptions::SaveExportedFonts()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // Configure a SaveOptions object to export fonts to separate files.
    // Set a callback that will handle font saving in a custom manner.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportFontResources(true);
    options->set_FontSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExHtmlSaveOptions::HandleFontSaving>());
    
    // The callback will export .ttf files and save them alongside the output document.
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.SaveExportedFonts.html", options);
    
    for (System::String fontFilename : System::Array<System::String>::FindAll(System::IO::Directory::GetFiles(get_ArtifactsDir()), static_cast<std::function<bool(System::String s)>>([](System::String s) -> bool
    {
        return s.EndsWith(u".ttf");
    })))
    {
        std::cout << fontFilename << std::endl;
    }
    
    
    ASSERT_EQ(10, System::Array<System::String>::FindAll(System::IO::Directory::GetFiles(get_ArtifactsDir()), static_cast<std::function<bool(System::String s)>>([](System::String s) -> bool
    {
        return s.EndsWith(u".ttf");
    }))->get_Length());
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, SaveExportedFonts)
{
    s_instance->SaveExportedFonts();
}

} // namespace gtest_test

void ExHtmlSaveOptions::HtmlVersions(Aspose::Words::Saving::HtmlVersion htmlVersion)
{
    //ExStart
    //ExFor:HtmlSaveOptions.#ctor(SaveFormat)
    //ExFor:HtmlSaveOptions.HtmlVersion
    //ExFor:HtmlVersion
    //ExSummary:Shows how to save a document to a specific version of HTML.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Html);
    options->set_HtmlVersion(htmlVersion);
    options->set_PrettyFormat(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.HtmlVersions.html", options);
    
    // Our HTML documents will have minor differences to be compatible with different HTML versions.
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.HtmlVersions.html");
    
    switch (htmlVersion)
    {
        case Aspose::Words::Saving::HtmlVersion::Html5:
            ASSERT_TRUE(outDocContents.Contains(u"<a id=\"_Toc76372689\"></a>"));
            ASSERT_TRUE(outDocContents.Contains(u"<a id=\"_Toc76372689\"></a>"));
            ASSERT_TRUE(outDocContents.Contains(u"<table style=\"padding:0pt; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
            break;
        
        case Aspose::Words::Saving::HtmlVersion::Xhtml:
            ASSERT_TRUE(outDocContents.Contains(u"<a name=\"_Toc76372689\"></a>"));
            ASSERT_TRUE(outDocContents.Contains(u"<ul type=\"disc\" style=\"margin:0pt; padding-left:0pt\">"));
            ASSERT_TRUE(outDocContents.Contains(u"<table cellspacing=\"0\" cellpadding=\"0\" style=\"-aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\""));
            break;
        
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_HtmlVersions_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::HtmlVersions)>::type;

struct ExHtmlSaveOptions_HtmlVersions : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_HtmlVersions_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::HtmlVersion::Html5),
            std::make_tuple(Aspose::Words::Saving::HtmlVersion::Xhtml),
        };
    }
};

TEST_P(ExHtmlSaveOptions_HtmlVersions, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->HtmlVersions(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_HtmlVersions, ::testing::ValuesIn(ExHtmlSaveOptions_HtmlVersions::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ExportXhtmlTransitional(bool showDoctypeDeclaration)
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportXhtmlTransitional
    //ExFor:HtmlSaveOptions.HtmlVersion
    //ExFor:HtmlVersion
    //ExSummary:Shows how to display a DOCTYPE heading when converting documents to the Xhtml 1.0 transitional standard.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Hello world!");
    
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Html);
    options->set_HtmlVersion(Aspose::Words::Saving::HtmlVersion::Xhtml);
    options->set_ExportXhtmlTransitional(showDoctypeDeclaration);
    options->set_PrettyFormat(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportXhtmlTransitional.html", options);
    
    // Our document will only contain a DOCTYPE declaration heading if we have set the "ExportXhtmlTransitional" flag to "true".
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.ExportXhtmlTransitional.html");
    System::String newLine = System::Environment::get_NewLine();
    
    if (showDoctypeDeclaration)
    {
        ASSERT_TRUE(outDocContents.Contains(System::String::Format(u"<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"no\"?>{0}", newLine) + System::String::Format(u"<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">{0}", newLine) + u"<html xmlns=\"http://www.w3.org/1999/xhtml\">"));
    }
    else
    {
        ASSERT_TRUE(outDocContents.Contains(u"<html>"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_ExportXhtmlTransitional_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ExportXhtmlTransitional)>::type;

struct ExHtmlSaveOptions_ExportXhtmlTransitional : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportXhtmlTransitional_Args>
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

TEST_P(ExHtmlSaveOptions_ExportXhtmlTransitional, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportXhtmlTransitional(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportXhtmlTransitional, ::testing::ValuesIn(ExHtmlSaveOptions_ExportXhtmlTransitional::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::Doc2EpubSaveOptions()
{
    //ExStart
    //ExFor:DocumentSplitCriteria
    //ExFor:HtmlSaveOptions
    //ExFor:HtmlSaveOptions.#ctor
    //ExFor:HtmlSaveOptions.Encoding
    //ExFor:HtmlSaveOptions.DocumentSplitCriteria
    //ExFor:HtmlSaveOptions.ExportDocumentProperties
    //ExFor:HtmlSaveOptions.SaveFormat
    //ExFor:SaveOptions
    //ExFor:SaveOptions.SaveFormat
    //ExSummary:Shows how to use a specific encoding when saving a document to .epub.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // Use a SaveOptions object to specify the encoding for a document that we will save.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_SaveFormat(Aspose::Words::SaveFormat::Epub);
    saveOptions->set_Encoding(System::Text::Encoding::get_UTF8());
    
    // By default, an output .epub document will have all its contents in one HTML part.
    // A split criterion allows us to segment the document into several HTML parts.
    // We will set the criteria to split the document into heading paragraphs.
    // This is useful for readers who cannot read HTML files more significant than a specific size.
    saveOptions->set_DocumentSplitCriteria(Aspose::Words::Saving::DocumentSplitCriteria::HeadingParagraph);
    
    // Specify that we want to export document properties.
    saveOptions->set_ExportDocumentProperties(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.Doc2EpubSaveOptions.epub", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, Doc2EpubSaveOptions)
{
    s_instance->Doc2EpubSaveOptions();
}

} // namespace gtest_test

void ExHtmlSaveOptions::ContentIdUrls(bool exportCidUrlsForMhtmlResources)
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportCidUrlsForMhtmlResources
    //ExSummary:Shows how to enable content IDs for output MHTML documents.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // Setting this flag will replace "Content-Location" tags
    // with "Content-ID" tags for each resource from the input document.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Mhtml);
    options->set_ExportCidUrlsForMhtmlResources(exportCidUrlsForMhtmlResources);
    options->set_CssStyleSheetType(Aspose::Words::Saving::CssStyleSheetType::External);
    options->set_ExportFontResources(true);
    options->set_PrettyFormat(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ContentIdUrls.mht", options);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.ContentIdUrls.mht");
    
    if (exportCidUrlsForMhtmlResources)
    {
        ASSERT_TRUE(outDocContents.Contains(u"Content-ID: <document.html>"));
        ASSERT_TRUE(outDocContents.Contains(u"<link href=3D\"cid:styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />"));
        ASSERT_TRUE(outDocContents.Contains(u"@font-face { font-family:'Arial Black'; font-weight:bold; src:url('cid:arib=\r\nlk.ttf') }"));
        ASSERT_TRUE(outDocContents.Contains(u"<img src=3D\"cid:image.003.jpeg\" width=3D\"350\" height=3D\"180\" alt=3D\"\" />"));
    }
    else
    {
        ASSERT_TRUE(outDocContents.Contains(u"Content-Location: document.html"));
        ASSERT_TRUE(outDocContents.Contains(u"<link href=3D\"styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />"));
        ASSERT_TRUE(outDocContents.Contains(u"@font-face { font-family:'Arial Black'; font-weight:bold; src:url('ariblk.t=\r\ntf') }"));
        ASSERT_TRUE(outDocContents.Contains(u"<img src=3D\"image.003.jpeg\" width=3D\"350\" height=3D\"180\" alt=3D\"\" />"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_ContentIdUrls_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ContentIdUrls)>::type;

struct ExHtmlSaveOptions_ContentIdUrls : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ContentIdUrls_Args>
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

TEST_P(ExHtmlSaveOptions_ContentIdUrls, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ContentIdUrls(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ContentIdUrls, ::testing::ValuesIn(ExHtmlSaveOptions_ContentIdUrls::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::DropDownFormField(bool exportDropDownFormFieldAsText)
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportDropDownFormFieldAsText
    //ExSummary:Shows how to get drop-down combo box form fields to blend in with paragraph text when saving to html.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Use a document builder to insert a combo box with the value "Two" selected.
    builder->InsertComboBox(u"MyComboBox", System::MakeArray<System::String>({u"One", u"Two", u"Three"}), 1);
    
    // The "ExportDropDownFormFieldAsText" flag of this SaveOptions object allows us to
    // control how saving the document to HTML treats drop-down combo boxes.
    // Setting it to "true" will convert each combo box into simple text
    // that displays the combo box's currently selected value, effectively freezing it.
    // Setting it to "false" will preserve the functionality of the combo box using <select> and <option> tags.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportDropDownFormFieldAsText(exportDropDownFormFieldAsText);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.DropDownFormField.html", options);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.DropDownFormField.html");
    
    if (exportDropDownFormFieldAsText)
    {
        ASSERT_TRUE(outDocContents.Contains(u"<span>Two</span>"));
    }
    else
    {
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<select name=\"MyComboBox\">") + u"<option>One</option>" + u"<option selected=\"selected\">Two</option>" + u"<option>Three</option>" + u"</select>"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_DropDownFormField_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::DropDownFormField)>::type;

struct ExHtmlSaveOptions_DropDownFormField : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_DropDownFormField_Args>
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

TEST_P(ExHtmlSaveOptions_DropDownFormField, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DropDownFormField(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_DropDownFormField, ::testing::ValuesIn(ExHtmlSaveOptions_DropDownFormField::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ExportImagesAsBase64(bool exportImagesAsBase64)
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportFontsAsBase64
    //ExFor:HtmlSaveOptions.ExportImagesAsBase64
    //ExSummary:Shows how to save a .html document with images embedded inside it.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportImagesAsBase64(exportImagesAsBase64);
    options->set_PrettyFormat(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportImagesAsBase64.html", options);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.ExportImagesAsBase64.html");
    
    ASSERT_TRUE(exportImagesAsBase64 ? outDocContents.Contains(u"<img src=\"data:image/png;base64") : outDocContents.Contains(u"<img src=\"HtmlSaveOptions.ExportImagesAsBase64.001.png\""));
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_ExportImagesAsBase64_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ExportImagesAsBase64)>::type;

struct ExHtmlSaveOptions_ExportImagesAsBase64 : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportImagesAsBase64_Args>
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

TEST_P(ExHtmlSaveOptions_ExportImagesAsBase64, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportImagesAsBase64(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportImagesAsBase64, ::testing::ValuesIn(ExHtmlSaveOptions_ExportImagesAsBase64::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ExportFontsAsBase64()
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportFontsAsBase64
    //ExFor:HtmlSaveOptions.ExportImagesAsBase64
    //ExSummary:Shows how to embed fonts inside a saved HTML document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportFontsAsBase64(true);
    options->set_CssStyleSheetType(Aspose::Words::Saving::CssStyleSheetType::Embedded);
    options->set_PrettyFormat(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportFontsAsBase64.html", options);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, ExportFontsAsBase64)
{
    s_instance->ExportFontsAsBase64();
}

} // namespace gtest_test

void ExHtmlSaveOptions::ExportLanguageInformation(bool exportLanguageInformation)
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportLanguageInformation
    //ExSummary:Shows how to preserve language information when saving to .html.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Use the builder to write text while formatting it in different locales.
    builder->get_Font()->set_LocaleId(System::MakeObject<System::Globalization::CultureInfo>(u"en-US")->get_LCID());
    builder->Writeln(u"Hello world!");
    
    builder->get_Font()->set_LocaleId(System::MakeObject<System::Globalization::CultureInfo>(u"en-GB")->get_LCID());
    builder->Writeln(u"Hello again!");
    
    builder->get_Font()->set_LocaleId(System::MakeObject<System::Globalization::CultureInfo>(u"ru-RU")->get_LCID());
    builder->Write(u"Привет, мир!");
    
    // When saving the document to HTML, we can pass a SaveOptions object
    // to either preserve or discard each formatted text's locale.
    // If we set the "ExportLanguageInformation" flag to "true",
    // the output HTML document will contain the locales in "lang" attributes of <span> tags.
    // If we set the "ExportLanguageInformation" flag to "false',
    // the text in the output HTML document will not contain any locale information.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportLanguageInformation(exportLanguageInformation);
    options->set_PrettyFormat(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportLanguageInformation.html", options);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.ExportLanguageInformation.html");
    
    if (exportLanguageInformation)
    {
        ASSERT_TRUE(outDocContents.Contains(u"<span>Hello world!</span>"));
        ASSERT_TRUE(outDocContents.Contains(u"<span lang=\"en-GB\">Hello again!</span>"));
        ASSERT_TRUE(outDocContents.Contains(u"<span lang=\"ru-RU\">Привет, мир!</span>"));
    }
    else
    {
        ASSERT_TRUE(outDocContents.Contains(u"<span>Hello world!</span>"));
        ASSERT_TRUE(outDocContents.Contains(u"<span>Hello again!</span>"));
        ASSERT_TRUE(outDocContents.Contains(u"<span>Привет, мир!</span>"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_ExportLanguageInformation_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ExportLanguageInformation)>::type;

struct ExHtmlSaveOptions_ExportLanguageInformation : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportLanguageInformation_Args>
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

TEST_P(ExHtmlSaveOptions_ExportLanguageInformation, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportLanguageInformation(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportLanguageInformation, ::testing::ValuesIn(ExHtmlSaveOptions_ExportLanguageInformation::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::List(Aspose::Words::Saving::ExportListLabels exportListLabels)
{
    //ExStart
    //ExFor:ExportListLabels
    //ExFor:HtmlSaveOptions.ExportListLabels
    //ExSummary:Shows how to configure list exporting to HTML.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::NumberDefault);
    builder->get_ListFormat()->set_List(list);
    
    builder->Writeln(u"Default numbered list item 1.");
    builder->Writeln(u"Default numbered list item 2.");
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"Default numbered list item 3.");
    builder->get_ListFormat()->RemoveNumbers();
    
    list = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::OutlineHeadingsLegal);
    builder->get_ListFormat()->set_List(list);
    
    builder->Writeln(u"Outline legal heading list item 1.");
    builder->Writeln(u"Outline legal heading list item 2.");
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"Outline legal heading list item 3.");
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"Outline legal heading list item 4.");
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"Outline legal heading list item 5.");
    builder->get_ListFormat()->RemoveNumbers();
    
    // When saving the document to HTML, we can pass a SaveOptions object
    // to decide which HTML elements the document will use to represent lists.
    // Setting the "ExportListLabels" property to "ExportListLabels.AsInlineText"
    // will create lists by formatting spans.
    // Setting the "ExportListLabels" property to "ExportListLabels.Auto" will use the <p> tag
    // to build lists in cases when using the <ol> and <li> tags may cause loss of formatting.
    // Setting the "ExportListLabels" property to "ExportListLabels.ByHtmlTags"
    // will use <ol> and <li> tags to build all lists.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportListLabels(exportListLabels);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.List.html", options);
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.List.html");
    
    switch (exportListLabels)
    {
        case Aspose::Words::Saving::ExportListLabels::AsInlineText:
            ASSERT_TRUE(outDocContents.Contains(System::String(u"<p style=\"margin-top:0pt; margin-left:72pt; margin-bottom:0pt; text-indent:-18pt; -aw-import:list-item; -aw-list-level-number:1; -aw-list-number-format:'%1.'; -aw-list-number-styles:'lowerLetter'; -aw-list-number-values:'1'; -aw-list-padding-sml:9.67pt\">") + u"<span style=\"-aw-import:ignore\">" + u"<span>a.</span>" + u"<span style=\"width:9.67pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" + u"</span>" + u"<span>Default numbered list item 3.</span>" + u"</p>"));
            ASSERT_TRUE(outDocContents.Contains(System::String(u"<p style=\"margin-top:0pt; margin-left:43.2pt; margin-bottom:0pt; text-indent:-43.2pt; -aw-import:list-item; -aw-list-level-number:3; -aw-list-number-format:'%0.%1.%2.%3'; -aw-list-number-styles:'decimal decimal decimal decimal'; -aw-list-number-values:'2 1 1 1'; -aw-list-padding-sml:10.2pt\">") + u"<span style=\"-aw-import:ignore\">" + u"<span>2.1.1.1</span>" + u"<span style=\"width:10.2pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" + u"</span>" + u"<span>Outline legal heading list item 5.</span>" + u"</p>"));
            break;
        
        case Aspose::Words::Saving::ExportListLabels::Auto:
            ASSERT_TRUE(outDocContents.Contains(System::String(u"<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">") + u"<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" + u"<span>Default numbered list item 3.</span>" + u"</li>" + u"</ol>"));
            break;
        
        case Aspose::Words::Saving::ExportListLabels::ByHtmlTags:
            ASSERT_TRUE(outDocContents.Contains(System::String(u"<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">") + u"<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" + u"<span>Default numbered list item 3.</span>" + u"</li>" + u"</ol>"));
            break;
        
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_List_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::List)>::type;

struct ExHtmlSaveOptions_List : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_List_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::ExportListLabels::AsInlineText),
            std::make_tuple(Aspose::Words::Saving::ExportListLabels::Auto),
            std::make_tuple(Aspose::Words::Saving::ExportListLabels::ByHtmlTags),
        };
    }
};

TEST_P(ExHtmlSaveOptions_List, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->List(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_List, ::testing::ValuesIn(ExHtmlSaveOptions_List::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ExportPageMargins(bool exportPageMargins)
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportPageMargins
    //ExSummary:Shows how to show out-of-bounds objects in output HTML documents.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Use a builder to insert a shape with no wrapping.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertShape(Aspose::Words::Drawing::ShapeType::Cube, 200, 200);
    
    shape->set_RelativeHorizontalPosition(Aspose::Words::Drawing::RelativeHorizontalPosition::Page);
    shape->set_RelativeVerticalPosition(Aspose::Words::Drawing::RelativeVerticalPosition::Page);
    shape->set_WrapType(Aspose::Words::Drawing::WrapType::None);
    
    // Negative shape position values may place the shape outside of page boundaries.
    // If we export this to HTML, the shape will appear truncated.
    shape->set_Left(-150);
    
    // When saving the document to HTML, we can pass a SaveOptions object
    // to decide whether to adjust the page to display out-of-bounds objects fully.
    // If we set the "ExportPageMargins" flag to "true", the shape will be fully visible in the output HTML.
    // If we set the "ExportPageMargins" flag to "false",
    // our document will display the shape truncated as we would see it in Microsoft Word.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportPageMargins(exportPageMargins);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportPageMargins.html", options);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.ExportPageMargins.html");
    
    if (exportPageMargins)
    {
        ASSERT_TRUE(outDocContents.Contains(u"<style type=\"text/css\">div.Section_1 { margin:70.85pt }</style>"));
        ASSERT_TRUE(outDocContents.Contains(u"<div class=\"Section_1\"><p style=\"margin-top:0pt; margin-left:150pt; margin-bottom:0pt\">"));
    }
    else
    {
        ASSERT_FALSE(outDocContents.Contains(u"style type=\"text/css\">"));
        ASSERT_TRUE(outDocContents.Contains(u"<div><p style=\"margin-top:0pt; margin-left:220.85pt; margin-bottom:0pt\">"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_ExportPageMargins_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ExportPageMargins)>::type;

struct ExHtmlSaveOptions_ExportPageMargins : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportPageMargins_Args>
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

TEST_P(ExHtmlSaveOptions_ExportPageMargins, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportPageMargins(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportPageMargins, ::testing::ValuesIn(ExHtmlSaveOptions_ExportPageMargins::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ExportPageSetup(bool exportPageSetup)
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportPageSetup
    //ExSummary:Shows how decide whether to preserve section structure/page setup information when saving to HTML.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Section 1");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Section 2");
    
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
    pageSetup->set_TopMargin(36.0);
    pageSetup->set_BottomMargin(36.0);
    pageSetup->set_PaperSize(Aspose::Words::PaperSize::A5);
    
    // When saving the document to HTML, we can pass a SaveOptions object
    // to decide whether to preserve or discard page setup settings.
    // If we set the "ExportPageSetup" flag to "true", the output HTML document will contain our page setup configuration.
    // If we set the "ExportPageSetup" flag to "false", the save operation will discard our page setup settings
    // for the first section, and both sections will look identical.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportPageSetup(exportPageSetup);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportPageSetup.html", options);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.ExportPageSetup.html");
    
    if (exportPageSetup)
    {
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<style type=\"text/css\">") + u"@page Section_1 { size:419.55pt 595.3pt; margin:36pt 70.85pt; -aw-footer-distance:35.4pt; -aw-header-distance:35.4pt }" + u"@page Section_2 { size:612pt 792pt; margin:70.85pt; -aw-footer-distance:35.4pt; -aw-header-distance:35.4pt }" + u"div.Section_1 { page:Section_1 }div.Section_2 { page:Section_2 }" + u"</style>"));
        
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<div class=\"Section_1\">") + u"<p style=\"margin-top:0pt; margin-bottom:0pt\">" + u"<span>Section 1</span>" + u"</p>" + u"</div>"));
    }
    else
    {
        ASSERT_FALSE(outDocContents.Contains(u"style type=\"text/css\">"));
        
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<div>") + u"<p style=\"margin-top:0pt; margin-bottom:0pt\">" + u"<span>Section 1</span>" + u"</p>" + u"</div>"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_ExportPageSetup_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ExportPageSetup)>::type;

struct ExHtmlSaveOptions_ExportPageSetup : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportPageSetup_Args>
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

TEST_P(ExHtmlSaveOptions_ExportPageSetup, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportPageSetup(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportPageSetup, ::testing::ValuesIn(ExHtmlSaveOptions_ExportPageSetup::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::RelativeFontSize(bool exportRelativeFontSize)
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportRelativeFontSize
    //ExSummary:Shows how to use relative font sizes when saving to .html.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Default font size, ");
    builder->get_Font()->set_Size(24);
    builder->Writeln(u"2x default font size,");
    builder->get_Font()->set_Size(96);
    builder->Write(u"8x default font size");
    
    // When we save the document to HTML, we can pass a SaveOptions object
    // to determine whether to use relative or absolute font sizes.
    // Set the "ExportRelativeFontSize" flag to "true" to declare font sizes
    // using the "em" measurement unit, which is a factor that multiplies the current font size. 
    // Set the "ExportRelativeFontSize" flag to "false" to declare font sizes
    // using the "pt" measurement unit, which is the font's absolute size in points.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportRelativeFontSize(exportRelativeFontSize);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.RelativeFontSize.html", options);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.RelativeFontSize.html");
    
    if (exportRelativeFontSize)
    {
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<body style=\"font-family:'Times New Roman'\">") + u"<div>" + u"<p style=\"margin-top:0pt; margin-bottom:0pt\">" + u"<span>Default font size, </span>" + u"</p>" + u"<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:2em\">" + u"<span>2x default font size,</span>" + u"</p>" + u"<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:8em\">" + u"<span>8x default font size</span>" + u"</p>" + u"</div>" + u"</body>"));
    }
    else
    {
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<body style=\"font-family:'Times New Roman'; font-size:12pt\">") + u"<div>" + u"<p style=\"margin-top:0pt; margin-bottom:0pt\">" + u"<span>Default font size, </span>" + u"</p>" + u"<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:24pt\">" + u"<span>2x default font size,</span>" + u"</p>" + u"<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:96pt\">" + u"<span>8x default font size</span>" + u"</p>" + u"</div>" + u"</body>"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_RelativeFontSize_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::RelativeFontSize)>::type;

struct ExHtmlSaveOptions_RelativeFontSize : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_RelativeFontSize_Args>
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

TEST_P(ExHtmlSaveOptions_RelativeFontSize, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RelativeFontSize(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_RelativeFontSize, ::testing::ValuesIn(ExHtmlSaveOptions_RelativeFontSize::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ExportShape(bool exportShapesAsSvg)
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportShapesAsSvg
    //ExSummary:Shows how to export shape as scalable vector graphics.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBox = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100.0, 60.0);
    builder->MoveTo(textBox->get_FirstParagraph());
    builder->Write(u"My text box");
    
    // When we save the document to HTML, we can pass a SaveOptions object
    // to determine how the saving operation will export text box shapes.
    // If we set the "ExportTextBoxAsSvg" flag to "true",
    // the save operation will convert shapes with text into SVG objects.
    // If we set the "ExportTextBoxAsSvg" flag to "false",
    // the save operation will convert shapes with text into images.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportShapesAsSvg(exportShapesAsSvg);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportTextBox.html", options);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.ExportTextBox.html");
    
    if (exportShapesAsSvg)
    {
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">") + u"<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"133\" height=\"80\">"));
    }
    else
    {
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<p style=\"margin-top:0pt; margin-bottom:0pt\">") + u"<img src=\"HtmlSaveOptions.ExportTextBox.001.png\" width=\"136\" height=\"83\" alt=\"\" " + u"style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" + u"</p>"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_ExportShape_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ExportShape)>::type;

struct ExHtmlSaveOptions_ExportShape : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportShape_Args>
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

TEST_P(ExHtmlSaveOptions_ExportShape, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportShape(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportShape, ::testing::ValuesIn(ExHtmlSaveOptions_ExportShape::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::RoundTripInformation(bool exportRoundtripInformation)
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportRoundtripInformation
    //ExSummary:Shows how to preserve hidden elements when converting to .html.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // When converting a document to .html, some elements such as hidden bookmarks, original shape positions,
    // or footnotes will be either removed or converted to plain text and effectively be lost.
    // Saving with a HtmlSaveOptions object with ExportRoundtripInformation set to true will preserve these elements.
    
    // When we save the document to HTML, we can pass a SaveOptions object to determine
    // how the saving operation will export document elements that HTML does not support or use,
    // such as hidden bookmarks and original shape positions.
    // If we set the "ExportRoundtripInformation" flag to "true", the save operation will preserve these elements.
    // If we set the "ExportRoundTripInformation" flag to "false", the save operation will discard these elements.
    // We will want to preserve such elements if we intend to load the saved HTML using Aspose.Words,
    // as they could be of use once again.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportRoundtripInformation(exportRoundtripInformation);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.RoundTripInformation.html", options);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.RoundTripInformation.html");
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"HtmlSaveOptions.RoundTripInformation.html");
    
    if (exportRoundtripInformation)
    {
        ASSERT_TRUE(outDocContents.Contains(u"<div style=\"-aw-headerfooter-type:header-primary; clear:both\">"));
        ASSERT_TRUE(outDocContents.Contains(u"<span style=\"-aw-import:ignore\">&#xa0;</span>"));
        
        ASSERT_TRUE(outDocContents.Contains(System::String(u"td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; ") + u"padding-right:2.4pt; padding-left:5.03pt; vertical-align:top; " + u"-aw-border-bottom:0.5pt single; -aw-border-left:0.5pt single; -aw-border-top:0.5pt single\">"));
        
        ASSERT_TRUE(outDocContents.Contains(u"<li style=\"margin-left:30.2pt; padding-left:5.8pt; -aw-font-family:'Courier New'; -aw-font-weight:normal; -aw-number-format:'o'\">"));
        
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"350\" height=\"180\" alt=\"\" ") + u"style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />"));
        
        
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<span>Page number </span>") + u"<span style=\"-aw-field-start:true\"></span>" + u"<span style=\"-aw-field-code:' PAGE   \\\\* MERGEFORMAT '\"></span>" + u"<span style=\"-aw-field-separator:true\"></span>" + u"<span>1</span>" + u"<span style=\"-aw-field-end:true\"></span>"));
        
        ASSERT_EQ(1, doc->get_Range()->get_Fields()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fields::Field>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fields::Field> f)>>([](System::SharedPtr<Aspose::Words::Fields::Field> f) -> bool
        {
            return f->get_Type() == Aspose::Words::Fields::FieldType::FieldPage;
        }))));
    }
    else
    {
        ASSERT_TRUE(outDocContents.Contains(u"<div style=\"clear:both\">"));
        ASSERT_TRUE(outDocContents.Contains(u"<span>&#xa0;</span>"));
        
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; ") + u"padding-right:2.4pt; padding-left:5.03pt; vertical-align:top\">"));
        
        ASSERT_TRUE(outDocContents.Contains(u"<li style=\"margin-left:30.2pt; padding-left:5.8pt\">"));
        
        ASSERT_TRUE(outDocContents.Contains(u"<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"350\" height=\"180\" alt=\"\" />"));
        
        ASSERT_TRUE(outDocContents.Contains(u"<span>Page number 1</span>"));
        
        ASSERT_EQ(0, doc->get_Range()->get_Fields()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fields::Field>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fields::Field> f)>>([](System::SharedPtr<Aspose::Words::Fields::Field> f) -> bool
        {
            return f->get_Type() == Aspose::Words::Fields::FieldType::FieldPage;
        }))));
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_RoundTripInformation_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::RoundTripInformation)>::type;

struct ExHtmlSaveOptions_RoundTripInformation : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_RoundTripInformation_Args>
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

TEST_P(ExHtmlSaveOptions_RoundTripInformation, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RoundTripInformation(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_RoundTripInformation, ::testing::ValuesIn(ExHtmlSaveOptions_RoundTripInformation::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ExportTocPageNumbers(bool exportTocPageNumbers)
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportTocPageNumbers
    //ExSummary:Shows how to display page numbers when saving a document with a table of contents to .html.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a table of contents, and then populate the document with paragraphs formatted using a "Heading"
    // style that the table of contents will pick up as entries. Each entry will display the heading paragraph on the left,
    // and the page number that contains the heading on the right.
    auto fieldToc = System::ExplicitCast<Aspose::Words::Fields::FieldToc>(builder->InsertField(Aspose::Words::Fields::FieldType::FieldTOC, true));
    
    builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(u"Heading 1"));
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Entry 1");
    builder->Writeln(u"Entry 2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Entry 3");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Entry 4");
    fieldToc->UpdatePageNumbers();
    doc->UpdateFields();
    
    // HTML documents do not have pages. If we save this document to HTML,
    // the page numbers that our TOC displays will have no meaning.
    // When we save the document to HTML, we can pass a SaveOptions object to omit these page numbers from the TOC.
    // If we set the "ExportTocPageNumbers" flag to "true",
    // each TOC entry will display the heading, separator, and page number, preserving its appearance in Microsoft Word.
    // If we set the "ExportTocPageNumbers" flag to "false",
    // the save operation will omit both the separator and page number and leave the heading for each entry intact.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportTocPageNumbers(exportTocPageNumbers);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ExportTocPageNumbers.html", options);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.ExportTocPageNumbers.html");
    
    if (exportTocPageNumbers)
    {
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<span>Entry 1</span>") + u"<span style=\"width:428.14pt; font-family:'Lucida Console'; font-size:10pt; display:inline-block; -aw-font-family:'Times New Roman'; " + u"-aw-tabstop-align:right; -aw-tabstop-leader:dots; -aw-tabstop-pos:469.8pt\">.......................................................................</span>" + u"<span>2</span>" + u"</p>"));
    }
    else
    {
        ASSERT_TRUE(outDocContents.Contains(System::String(u"<p style=\"margin-top:0pt; margin-bottom:0pt\">") + u"<span>Entry 2</span>" + u"</p>"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_ExportTocPageNumbers_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ExportTocPageNumbers)>::type;

struct ExHtmlSaveOptions_ExportTocPageNumbers : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ExportTocPageNumbers_Args>
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

TEST_P(ExHtmlSaveOptions_ExportTocPageNumbers, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportTocPageNumbers(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ExportTocPageNumbers, ::testing::ValuesIn(ExHtmlSaveOptions_ExportTocPageNumbers::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::FontSubsetting(int32_t fontResourcesSubsettingSizeThreshold)
{
    //ExStart
    //ExFor:HtmlSaveOptions.FontResourcesSubsettingSizeThreshold
    //ExSummary:Shows how to work with font subsetting.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Arial");
    builder->Writeln(u"Hello world!");
    builder->get_Font()->set_Name(u"Times New Roman");
    builder->Writeln(u"Hello world!");
    builder->get_Font()->set_Name(u"Courier New");
    builder->Writeln(u"Hello world!");
    
    // When we save the document to HTML, we can pass a SaveOptions object configure font subsetting.
    // Suppose we set the "ExportFontResources" flag to "true" and also name a folder in the "FontsFolder" property.
    // In that case, the saving operation will create that folder and place a .ttf file inside
    // that folder for each font that our document uses.
    // Each .ttf file will contain that font's entire glyph set,
    // which may potentially result in a very large file that accompanies the document.
    // When we apply subsetting to a font, its exported raw data will only contain the glyphs that the document is
    // using instead of the entire glyph set. If the text in our document only uses a small fraction of a font's
    // glyph set, then subsetting will significantly reduce our output documents' size.
    // We can use the "FontResourcesSubsettingSizeThreshold" property to define a .ttf file size, in bytes.
    // If an exported font creates a size bigger file than that, then the save operation will apply subsetting to that font. 
    // Setting a threshold of 0 applies subsetting to all fonts,
    // and setting it to "int.MaxValue" effectively disables subsetting.
    System::String fontsFolder = get_ArtifactsDir() + u"HtmlSaveOptions.FontSubsetting.Fonts";
    
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ExportFontResources(true);
    options->set_FontsFolder(fontsFolder);
    options->set_FontResourcesSubsettingSizeThreshold(fontResourcesSubsettingSizeThreshold);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.FontSubsetting.html", options);
    
    System::ArrayPtr<System::String> fontFileNames = System::IO::Directory::GetFiles(fontsFolder)->LINQ_Where(static_cast<System::Func<System::String, bool>>(static_cast<std::function<bool(System::String s)>>([](System::String s) -> bool
    {
        return s.EndsWith(u".ttf");
    })))->LINQ_ToArray();
    
    ASSERT_EQ(3, fontFileNames->get_Length());
    
    for (System::String filename : fontFileNames)
    {
        // By default, the .ttf files for each of our three fonts will be over 700MB.
        // Subsetting will reduce them all to under 30MB.
        auto fontFileInfo = System::MakeObject<System::IO::FileInfo>(filename);
        
        ASSERT_TRUE(fontFileInfo->get_Length() > 700000 || fontFileInfo->get_Length() < 30000);
        ASSERT_TRUE(System::Math::Max(fontResourcesSubsettingSizeThreshold, 30000) > System::MakeObject<System::IO::FileInfo>(filename)->get_Length());
    }
    
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_FontSubsetting_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::FontSubsetting)>::type;

struct ExHtmlSaveOptions_FontSubsetting : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_FontSubsetting_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(0),
            std::make_tuple(1000000),
            std::make_tuple(std::numeric_limits<int32_t>::max()),
        };
    }
};

TEST_P(ExHtmlSaveOptions_FontSubsetting, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FontSubsetting(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_FontSubsetting, ::testing::ValuesIn(ExHtmlSaveOptions_FontSubsetting::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::MetafileFormat(Aspose::Words::Saving::HtmlMetafileFormat htmlMetafileFormat)
{
    //ExStart
    //ExFor:HtmlMetafileFormat
    //ExFor:HtmlSaveOptions.MetafileFormat
    //ExFor:HtmlLoadOptions.ConvertSvgToEmf
    //ExSummary:Shows how to convert SVG objects to a different format when saving HTML documents.
    System::String html = u"<html>\r\n                    <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>\r\n                        <text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>\r\n                    </svg>\r\n                </html>";
    
    // Use 'ConvertSvgToEmf' to turn back the legacy behavior
    // where all SVG images loaded from an HTML document were converted to EMF.
    // Now SVG images are loaded without conversion
    // if the MS Word version specified in load options supports SVG images natively.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::HtmlLoadOptions>();
    loadOptions->set_ConvertSvgToEmf(true);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(System::MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(html)), loadOptions);
    
    // This document contains a <svg> element in the form of text.
    // When we save the document to HTML, we can pass a SaveOptions object
    // to determine how the saving operation handles this object.
    // Setting the "MetafileFormat" property to "HtmlMetafileFormat.Png" to convert it to a PNG image.
    // Setting the "MetafileFormat" property to "HtmlMetafileFormat.Svg" preserve it as a SVG object.
    // Setting the "MetafileFormat" property to "HtmlMetafileFormat.EmfOrWmf" to convert it to a metafile.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_MetafileFormat(htmlMetafileFormat);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.MetafileFormat.html", options);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.MetafileFormat.html");
    
    switch (htmlMetafileFormat)
    {
        case Aspose::Words::Saving::HtmlMetafileFormat::Png:
            ASSERT_TRUE(outDocContents.Contains(System::String(u"<p style=\"margin-top:0pt; margin-bottom:0pt\">") + u"<img src=\"HtmlSaveOptions.MetafileFormat.001.png\" width=\"500\" height=\"40\" alt=\"\" " + u"style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" + u"</p>"));
            break;
        
        case Aspose::Words::Saving::HtmlMetafileFormat::Svg:
            ASSERT_TRUE(outDocContents.Contains(System::String(u"<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">") + u"<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"499\" height=\"40\">"));
            break;
        
        case Aspose::Words::Saving::HtmlMetafileFormat::EmfOrWmf:
            ASSERT_TRUE(outDocContents.Contains(System::String(u"<p style=\"margin-top:0pt; margin-bottom:0pt\">") + u"<img src=\"HtmlSaveOptions.MetafileFormat.001.emf\" width=\"500\" height=\"40\" alt=\"\" " + u"style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" + u"</p>"));
            break;
        
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_MetafileFormat_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::MetafileFormat)>::type;

struct ExHtmlSaveOptions_MetafileFormat : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_MetafileFormat_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::HtmlMetafileFormat::Png),
            std::make_tuple(Aspose::Words::Saving::HtmlMetafileFormat::Svg),
            std::make_tuple(Aspose::Words::Saving::HtmlMetafileFormat::EmfOrWmf),
        };
    }
};

TEST_P(ExHtmlSaveOptions_MetafileFormat, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MetafileFormat(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_MetafileFormat, ::testing::ValuesIn(ExHtmlSaveOptions_MetafileFormat::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::OfficeMathOutputMode(Aspose::Words::Saving::HtmlOfficeMathOutputMode htmlOfficeMathOutputMode)
{
    //ExStart
    //ExFor:HtmlOfficeMathOutputMode
    //ExFor:HtmlSaveOptions.OfficeMathOutputMode
    //ExSummary:Shows how to specify how to export Microsoft OfficeMath objects to HTML.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    // When we save the document to HTML, we can pass a SaveOptions object
    // to determine how the saving operation handles OfficeMath objects.
    // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.Image"
    // will render each OfficeMath object into an image.
    // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.MathML"
    // will convert each OfficeMath object into MathML.
    // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.Text"
    // will represent each OfficeMath formula using plain HTML text.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_OfficeMathOutputMode(htmlOfficeMathOutputMode);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.OfficeMathOutputMode.html", options);
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.OfficeMathOutputMode.html");
    
    switch (htmlOfficeMathOutputMode)
    {
        case Aspose::Words::Saving::HtmlOfficeMathOutputMode::Image:
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, System::String(u"<p style=\"margin-top:0pt; margin-bottom:10pt\">") + u"<img src=\"HtmlSaveOptions.OfficeMathOutputMode.001.png\" width=\"163\" height=\"19\" alt=\"\" style=\"vertical-align:middle; " + u"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" + u"</p>")->get_Success());
            break;
        
        case Aspose::Words::Saving::HtmlOfficeMathOutputMode::MathML:
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, System::String(u"<p style=\"margin-top:0pt; margin-bottom:10pt; text-align:center\">") + u"<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" + u"<mi>i</mi>" + u"<mo>[+]</mo>" + u"<mi>b</mi>" + u"<mo>-</mo>" + u"<mi>c</mi>" + u"<mo>≥</mo>" + u".*" + u"</math>" + u"</p>")->get_Success());
            break;
        
        case Aspose::Words::Saving::HtmlOfficeMathOutputMode::Text:
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, System::String(u"<p style=\\\"margin-top:0pt; margin-bottom:10pt; text-align:center\\\">") + u"<span style=\\\"font-family:'Cambria Math'\\\">i[+]b-c≥iM[+]bM-cM </span>" + u"</p>")->get_Success());
            break;
        
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_OfficeMathOutputMode_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::OfficeMathOutputMode)>::type;

struct ExHtmlSaveOptions_OfficeMathOutputMode : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_OfficeMathOutputMode_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::HtmlOfficeMathOutputMode::Image),
            std::make_tuple(Aspose::Words::Saving::HtmlOfficeMathOutputMode::MathML),
            std::make_tuple(Aspose::Words::Saving::HtmlOfficeMathOutputMode::Text),
        };
    }
};

TEST_P(ExHtmlSaveOptions_OfficeMathOutputMode, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->OfficeMathOutputMode(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_OfficeMathOutputMode, ::testing::ValuesIn(ExHtmlSaveOptions_OfficeMathOutputMode::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ImageFolder()
{
    //ExStart
    //ExFor:HtmlSaveOptions
    //ExFor:HtmlSaveOptions.ExportTextInputFormFieldAsText
    //ExFor:HtmlSaveOptions.ImagesFolder
    //ExSummary:Shows how to specify the folder for storing linked images after saving to .html.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    System::String imagesDir = System::IO::Path::Combine(get_ArtifactsDir(), u"SaveHtmlWithOptions");
    
    if (System::IO::Directory::Exists(imagesDir))
    {
        System::IO::Directory::Delete(imagesDir, true);
    }
    
    System::IO::Directory::CreateDirectory_(imagesDir);
    
    // Set an option to export form fields as plain text instead of HTML input elements.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Html);
    options->set_ExportTextInputFormFieldAsText(true);
    options->set_ImagesFolder(imagesDir);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.SaveHtmlWithOptions.html", options);
    //ExEnd
    
    ASSERT_TRUE(System::IO::File::Exists(get_ArtifactsDir() + u"HtmlSaveOptions.SaveHtmlWithOptions.html"));
    ASSERT_EQ(9, System::IO::Directory::GetFiles(imagesDir)->get_Length());
    
    System::IO::Directory::Delete(imagesDir, true);
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, ImageFolder)
{
    s_instance->ImageFolder();
}

} // namespace gtest_test

void ExHtmlSaveOptions::ImageSavingCallback()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // When we save the document to HTML, we can pass a SaveOptions object to designate a callback
    // to customize the image saving process.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    options->set_ImageSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExHtmlSaveOptions::ImageShapePrinter>());
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ImageSavingCallback.html", options);
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, ImageSavingCallback)
{
    s_instance->ImageSavingCallback();
}

} // namespace gtest_test

void ExHtmlSaveOptions::PrettyFormat(bool usePrettyFormat)
{
    //ExStart
    //ExFor:SaveOptions.PrettyFormat
    //ExSummary:Shows how to enhance the readability of the raw code of a saved .html document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    auto htmlOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Html);
    htmlOptions->set_PrettyFormat(usePrettyFormat);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.PrettyFormat.html", htmlOptions);
    
    // Enabling pretty format makes the raw html code more readable by adding tab stop and new line characters.
    System::String html = System::IO::File::ReadAllText(get_ArtifactsDir() + u"HtmlSaveOptions.PrettyFormat.html");
    
    System::String newLine = System::Environment::get_NewLine();
    if (usePrettyFormat)
    {
        ASSERT_EQ(System::String::Format(u"<html>{0}", newLine) + System::String::Format(u"\t<head>{0}", newLine) + System::String::Format(u"\t\t<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />{0}", newLine) + System::String::Format(u"\t\t<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />{0}", newLine) + System::String::Format(u"\t\t<meta name=\"generator\" content=\"{0} {1}\" />{2}", Aspose::Words::BuildVersionInfo::get_Product(), Aspose::Words::BuildVersionInfo::get_Version(), newLine) + System::String::Format(u"\t\t<title>{0}", newLine) + System::String::Format(u"\t\t</title>{0}", newLine) + System::String::Format(u"\t</head>{0}", newLine) + System::String::Format(u"\t<body style=\"font-family:'Times New Roman'; font-size:12pt\">{0}", newLine) + System::String::Format(u"\t\t<div>{0}", newLine) + System::String::Format(u"\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">{0}", newLine) + System::String::Format(u"\t\t\t\t<span>Hello world!</span>{0}", newLine) + System::String::Format(u"\t\t\t</p>{0}", newLine) + System::String::Format(u"\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">{0}", newLine) + System::String::Format(u"\t\t\t\t<span style=\"-aw-import:ignore\">&#xa0;</span>{0}", newLine) + System::String::Format(u"\t\t\t</p>{0}", newLine) + System::String::Format(u"\t\t</div>{0}", newLine) + System::String::Format(u"\t</body>{0}</html>", newLine), html);
    }
    else
    {
        ASSERT_EQ(System::String(u"<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />") + u"<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />" + System::String::Format(u"<meta name=\"generator\" content=\"{0} {1}\" /><title></title></head>", Aspose::Words::BuildVersionInfo::get_Product(), Aspose::Words::BuildVersionInfo::get_Version()) + u"<body style=\"font-family:'Times New Roman'; font-size:12pt\">" + u"<div><p style=\"margin-top:0pt; margin-bottom:0pt\"><span>Hello world!</span></p>" + u"<p style=\"margin-top:0pt; margin-bottom:0pt\"><span style=\"-aw-import:ignore\">&#xa0;</span></p></div></body></html>", html);
    }
    //ExEnd
}

namespace gtest_test
{

using ExHtmlSaveOptions_PrettyFormat_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::PrettyFormat)>::type;

struct ExHtmlSaveOptions_PrettyFormat : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_PrettyFormat_Args>
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

TEST_P(ExHtmlSaveOptions_PrettyFormat, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PrettyFormat(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_PrettyFormat, ::testing::ValuesIn(ExHtmlSaveOptions_PrettyFormat::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::ProgressCallback(Aspose::Words::SaveFormat saveFormat, System::String ext)
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    
    // Following formats are supported: Html, Mhtml, Epub.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(saveFormat);
    saveOptions->set_ProgressCallback(System::MakeObject<Aspose::Words::ApiExamples::ExHtmlSaveOptions::SavingProgressCallback>());
    
    System::OperationCanceledException exception; ASPOSE_ASSERT_THROW(static_cast<std::function<void()>>([&doc, &ext, &saveOptions]() -> void
    {
        doc->Save(get_ArtifactsDir() + System::String::Format(u"HtmlSaveOptions.ProgressCallback.{0}", ext), saveOptions);
    })(), System::OperationCanceledException, &exception);
    System::Nullable<bool> actual = System::Default<System::Nullable<bool>>();
    System::OperationCanceledException condExpression = exception;
    if (condExpression != nullptr)
    {
        actual = condExpression->get_Message().Contains(u"EstimatedProgress");
    }
    ASSERT_TRUE(actual);
}

namespace gtest_test
{

using ExHtmlSaveOptions_ProgressCallback_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::ProgressCallback)>::type;

struct ExHtmlSaveOptions_ProgressCallback : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_ProgressCallback_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::SaveFormat::Html, u"html"),
            std::make_tuple(Aspose::Words::SaveFormat::Mhtml, u"mhtml"),
            std::make_tuple(Aspose::Words::SaveFormat::Epub, u"epub"),
        };
    }
};

TEST_P(ExHtmlSaveOptions_ProgressCallback, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ProgressCallback(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_ProgressCallback, ::testing::ValuesIn(ExHtmlSaveOptions_ProgressCallback::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::MobiAzw3DefaultEncoding(Aspose::Words::SaveFormat saveFormat)
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_SaveFormat(saveFormat);
    saveOptions->set_Encoding(System::Text::Encoding::get_ASCII());
    
    System::String outputFileName = System::String::Format(u"{0}HtmlSaveOptions.MobiDefaultEncoding{1}", get_ArtifactsDir(), Aspose::Words::FileFormatUtil::SaveFormatToExtension(saveFormat));
    doc->Save(outputFileName);
    
    System::SharedPtr<System::Text::Encoding> encoding = Aspose::Words::ApiExamples::TestUtil::GetEncoding(outputFileName);
    ASPOSE_ASSERT_NE(System::Text::Encoding::get_ASCII(), encoding);
    ASPOSE_ASSERT_EQ(System::Text::Encoding::get_UTF8(), encoding);
}

namespace gtest_test
{

using ExHtmlSaveOptions_MobiAzw3DefaultEncoding_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHtmlSaveOptions::MobiAzw3DefaultEncoding)>::type;

struct ExHtmlSaveOptions_MobiAzw3DefaultEncoding : public ExHtmlSaveOptions, public Aspose::Words::ApiExamples::ExHtmlSaveOptions, public ::testing::WithParamInterface<ExHtmlSaveOptions_MobiAzw3DefaultEncoding_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::SaveFormat::Mobi),
            std::make_tuple(Aspose::Words::SaveFormat::Azw3),
        };
    }
};

TEST_P(ExHtmlSaveOptions_MobiAzw3DefaultEncoding, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MobiAzw3DefaultEncoding(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHtmlSaveOptions_MobiAzw3DefaultEncoding, ::testing::ValuesIn(ExHtmlSaveOptions_MobiAzw3DefaultEncoding::TestCases()));

} // namespace gtest_test

void ExHtmlSaveOptions::HtmlReplaceBackslashWithYenSign()
{
    //ExStart:HtmlReplaceBackslashWithYenSign
    //GistId:708ce40a68fac5003d46f6b4acfd5ff1
    //ExFor:HtmlSaveOptions.ReplaceBackslashWithYenSign
    //ExSummary:Shows how to replace backslash characters with yen signs (Html).
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Korean backslash symbol.docx");
    
    // By default, Aspose.Words mimics MS Word's behavior and doesn't replace backslash characters with yen signs in
    // generated HTML documents. However, previous versions of Aspose.Words performed such replacements in certain
    // scenarios. This flag enables backward compatibility with previous versions of Aspose.Words.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_ReplaceBackslashWithYenSign(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.ReplaceBackslashWithYenSign.html", saveOptions);
    //ExEnd:HtmlReplaceBackslashWithYenSign
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, HtmlReplaceBackslashWithYenSign)
{
    s_instance->HtmlReplaceBackslashWithYenSign();
}

} // namespace gtest_test

void ExHtmlSaveOptions::RemoveJavaScriptFromLinks()
{
    //ExStart:HtmlRemoveJavaScriptFromLinks
    //GistId:12a3a3cfe30f3145220db88428a9f814
    //ExFor:HtmlFixedSaveOptions.RemoveJavaScriptFromLinks
    //ExSummary:Shows how to remove JavaScript from the links.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"JavaScript in HREF.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    saveOptions->set_RemoveJavaScriptFromLinks(true);
    
    doc->Save(get_ArtifactsDir() + u"HtmlSaveOptions.RemoveJavaScriptFromLinks.html", saveOptions);
    //ExEnd:HtmlRemoveJavaScriptFromLinks
}

namespace gtest_test
{

TEST_F(ExHtmlSaveOptions, RemoveJavaScriptFromLinks)
{
    s_instance->RemoveJavaScriptFromLinks();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
