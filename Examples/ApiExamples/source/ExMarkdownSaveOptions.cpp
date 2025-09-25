// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExMarkdownSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/path.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/func.h>
#include <system/environment.h>
#include <system/enum_helpers.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <functional>
#include <Aspose.Words.Cpp/Model/Text/Underline.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/MarkdownSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/MarkdownOfficeMathExportMode.h>
#include <Aspose.Words.Cpp/Model/Saving/MarkdownLinkExportMode.h>
#include <Aspose.Words.Cpp/Model/Saving/MarkdownExportAsHtml.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "DocumentHelper.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3457578932u, ::Aspose::Words::ApiExamples::ExMarkdownSaveOptions::SavedImageRename, ThisTypeBaseTypesInfo);

ExMarkdownSaveOptions::SavedImageRename::SavedImageRename(System::String outFileName) : mCount(0)
{
    mOutFileName = outFileName;
}

void ExMarkdownSaveOptions::SavedImageRename::ImageSaving(System::SharedPtr<Aspose::Words::Saving::ImageSavingArgs> args)
{
    System::String imageFileName = System::String::Format(u"{0} shape {1}, of type {2}{3}", mOutFileName, ++mCount, args->get_CurrentShape()->get_ShapeType(), System::IO::Path::GetExtension(args->get_ImageFileName()));
    
    args->set_ImageFileName(imageFileName);
    args->set_ImageStream(System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + imageFileName, System::IO::FileMode::Create));
    
    ASSERT_TRUE(args->get_ImageStream()->get_CanWrite());
    ASSERT_TRUE(args->get_IsImageAvailable());
    ASSERT_FALSE(args->get_KeepImageStreamOpen());
}


RTTI_INFO_IMPL_HASH(970531554u, ::Aspose::Words::ApiExamples::ExMarkdownSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExMarkdownSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExMarkdownSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExMarkdownSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExMarkdownSaveOptions> ExMarkdownSaveOptions::s_instance;

} // namespace gtest_test

void ExMarkdownSaveOptions::MarkdownDocumentTableContentAlignment(Aspose::Words::Saving::TableContentAlignment tableContentAlignment)
{
    //ExStart
    //ExFor:TableContentAlignment
    //ExFor:MarkdownSaveOptions.TableContentAlignment
    //ExSummary:Shows how to align contents in tables.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    builder->InsertCell();
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Right);
    builder->Write(u"Cell1");
    builder->InsertCell();
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    builder->Write(u"Cell2");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::MarkdownSaveOptions>();
    saveOptions->set_TableContentAlignment(tableContentAlignment);
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.MarkdownDocumentTableContentAlignment.md", saveOptions);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"MarkdownSaveOptions.MarkdownDocumentTableContentAlignment.md");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    switch (tableContentAlignment)
    {
        case Aspose::Words::Saving::TableContentAlignment::Auto:
            ASSERT_EQ(Aspose::Words::ParagraphAlignment::Right, table->get_FirstRow()->get_Cells()->idx_get(0)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, table->get_FirstRow()->get_Cells()->idx_get(1)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            break;
        
        case Aspose::Words::Saving::TableContentAlignment::Left:
            ASSERT_EQ(Aspose::Words::ParagraphAlignment::Left, table->get_FirstRow()->get_Cells()->idx_get(0)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            ASSERT_EQ(Aspose::Words::ParagraphAlignment::Left, table->get_FirstRow()->get_Cells()->idx_get(1)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            break;
        
        case Aspose::Words::Saving::TableContentAlignment::Center:
            ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, table->get_FirstRow()->get_Cells()->idx_get(0)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, table->get_FirstRow()->get_Cells()->idx_get(1)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            break;
        
        case Aspose::Words::Saving::TableContentAlignment::Right:
            ASSERT_EQ(Aspose::Words::ParagraphAlignment::Right, table->get_FirstRow()->get_Cells()->idx_get(0)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            ASSERT_EQ(Aspose::Words::ParagraphAlignment::Right, table->get_FirstRow()->get_Cells()->idx_get(1)->get_FirstParagraph()->get_ParagraphFormat()->get_Alignment());
            break;
        
    }
    //ExEnd
}

namespace gtest_test
{

using ExMarkdownSaveOptions_MarkdownDocumentTableContentAlignment_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExMarkdownSaveOptions::MarkdownDocumentTableContentAlignment)>::type;

struct ExMarkdownSaveOptions_MarkdownDocumentTableContentAlignment : public ExMarkdownSaveOptions, public Aspose::Words::ApiExamples::ExMarkdownSaveOptions, public ::testing::WithParamInterface<ExMarkdownSaveOptions_MarkdownDocumentTableContentAlignment_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::TableContentAlignment::Left),
            std::make_tuple(Aspose::Words::Saving::TableContentAlignment::Right),
            std::make_tuple(Aspose::Words::Saving::TableContentAlignment::Center),
            std::make_tuple(Aspose::Words::Saving::TableContentAlignment::Auto),
        };
    }
};

TEST_P(ExMarkdownSaveOptions_MarkdownDocumentTableContentAlignment, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MarkdownDocumentTableContentAlignment(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExMarkdownSaveOptions_MarkdownDocumentTableContentAlignment, ::testing::ValuesIn(ExMarkdownSaveOptions_MarkdownDocumentTableContentAlignment::TestCases()));

} // namespace gtest_test

void ExMarkdownSaveOptions::RenameImages()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::MarkdownSaveOptions>();
    // If we convert a document that contains images into Markdown, we will end up with one Markdown file which links to several images.
    // Each image will be in the form of a file in the local file system.
    // There is also a callback that can customize the name and file system location of each image.
    saveOptions->set_ImageSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExMarkdownSaveOptions::SavedImageRename>(u"MarkdownSaveOptions.HandleDocument.md"));
    saveOptions->set_SaveFormat(Aspose::Words::SaveFormat::Markdown);
    
    // The ImageSaving() method of our callback will be run at this time.
    doc->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.HandleDocument.md", saveOptions);
    
    ASSERT_EQ(1, System::IO::Directory::GetFiles(get_ArtifactsDir())->LINQ_Where(static_cast<System::Func<System::String, bool>>(static_cast<std::function<bool(System::String s)>>([](System::String s) -> bool
    {
        return s.StartsWith(get_ArtifactsDir() + u"MarkdownSaveOptions.HandleDocument.md shape");
    })))->LINQ_Count(static_cast<System::Func<System::String, bool>>(static_cast<std::function<bool(System::String f)>>([](System::String f) -> bool
    {
        return f.EndsWith(u".jpeg");
    }))));
    ASSERT_EQ(8, System::IO::Directory::GetFiles(get_ArtifactsDir())->LINQ_Where(static_cast<System::Func<System::String, bool>>(static_cast<std::function<bool(System::String s)>>([](System::String s) -> bool
    {
        return s.StartsWith(get_ArtifactsDir() + u"MarkdownSaveOptions.HandleDocument.md shape");
    })))->LINQ_Count(static_cast<System::Func<System::String, bool>>(static_cast<std::function<bool(System::String f)>>([](System::String f) -> bool
    {
        return f.EndsWith(u".png");
    }))));
}

namespace gtest_test
{

TEST_F(ExMarkdownSaveOptions, RenameImages)
{
    s_instance->RenameImages();
}

} // namespace gtest_test

void ExMarkdownSaveOptions::ExportImagesAsBase64(bool exportImagesAsBase64)
{
    //ExStart
    //ExFor:MarkdownSaveOptions.ExportImagesAsBase64
    //ExSummary:Shows how to save a .md document with images embedded inside it.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Images.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::MarkdownSaveOptions>();
    saveOptions->set_ExportImagesAsBase64(exportImagesAsBase64);
    
    doc->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.ExportImagesAsBase64.md", saveOptions);
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"MarkdownSaveOptions.ExportImagesAsBase64.md");
    
    ASSERT_TRUE(exportImagesAsBase64 ? outDocContents.Contains(u"data:image/jpeg;base64") : outDocContents.Contains(u"MarkdownSaveOptions.ExportImagesAsBase64.001.jpeg"));
    //ExEnd
}

namespace gtest_test
{

using ExMarkdownSaveOptions_ExportImagesAsBase64_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExMarkdownSaveOptions::ExportImagesAsBase64)>::type;

struct ExMarkdownSaveOptions_ExportImagesAsBase64 : public ExMarkdownSaveOptions, public Aspose::Words::ApiExamples::ExMarkdownSaveOptions, public ::testing::WithParamInterface<ExMarkdownSaveOptions_ExportImagesAsBase64_Args>
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

TEST_P(ExMarkdownSaveOptions_ExportImagesAsBase64, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportImagesAsBase64(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExMarkdownSaveOptions_ExportImagesAsBase64, ::testing::ValuesIn(ExMarkdownSaveOptions_ExportImagesAsBase64::TestCases()));

} // namespace gtest_test

void ExMarkdownSaveOptions::ListExportMode(Aspose::Words::Saving::MarkdownListExportMode markdownListExportMode)
{
    //ExStart
    //ExFor:MarkdownSaveOptions.ListExportMode
    //ExFor:MarkdownListExportMode
    //ExSummary:Shows how to list items will be written to the markdown document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"List item.docx");
    
    // Use MarkdownListExportMode.PlainText or MarkdownListExportMode.MarkdownSyntax to export list.
    auto options = System::MakeObject<Aspose::Words::Saving::MarkdownSaveOptions>();
    options->set_ListExportMode(markdownListExportMode);
    doc->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.ListExportMode.md", options);
    //ExEnd
}

namespace gtest_test
{

using ExMarkdownSaveOptions_ListExportMode_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExMarkdownSaveOptions::ListExportMode)>::type;

struct ExMarkdownSaveOptions_ListExportMode : public ExMarkdownSaveOptions, public Aspose::Words::ApiExamples::ExMarkdownSaveOptions, public ::testing::WithParamInterface<ExMarkdownSaveOptions_ListExportMode_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::MarkdownListExportMode::PlainText),
            std::make_tuple(Aspose::Words::Saving::MarkdownListExportMode::MarkdownSyntax),
        };
    }
};

TEST_P(ExMarkdownSaveOptions_ListExportMode, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ListExportMode(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExMarkdownSaveOptions_ListExportMode, ::testing::ValuesIn(ExMarkdownSaveOptions_ListExportMode::TestCases()));

} // namespace gtest_test

void ExMarkdownSaveOptions::ImagesFolder()
{
    //ExStart
    //ExFor:MarkdownSaveOptions.ImagesFolder
    //ExFor:MarkdownSaveOptions.ImagesFolderAlias
    //ExSummary:Shows how to specifies the name of the folder used to construct image URIs.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    builder->Writeln(u"Some image below:");
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    System::String imagesFolder = System::IO::Path::Combine(get_ArtifactsDir(), u"ImagesDir");
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::MarkdownSaveOptions>();
    // Use the "ImagesFolder" property to assign a folder in the local file system into which
    // Aspose.Words will save all the document's linked images.
    saveOptions->set_ImagesFolder(imagesFolder);
    // Use the "ImagesFolderAlias" property to use this folder
    // when constructing image URIs instead of the images folder's name.
    saveOptions->set_ImagesFolderAlias(u"http://example.com/images");
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.ImagesFolder.md", saveOptions);
    //ExEnd
    
    System::ArrayPtr<System::String> dirFiles = System::IO::Directory::GetFiles(imagesFolder, u"MarkdownSaveOptions.ImagesFolder.001.jpeg");
    ASSERT_EQ(1, dirFiles->get_Length());
    auto doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"MarkdownSaveOptions.ImagesFolder.md");
    doc->GetText().Contains(u"http://example.com/images/MarkdownSaveOptions.ImagesFolder.001.jpeg");
}

namespace gtest_test
{

TEST_F(ExMarkdownSaveOptions, ImagesFolder)
{
    s_instance->ImagesFolder();
}

} // namespace gtest_test

void ExMarkdownSaveOptions::ExportUnderlineFormatting()
{
    //ExStart:ExportUnderlineFormatting
    //GistId:eeeec1fbf118e95e7df3f346c91ed726
    //ExFor:MarkdownSaveOptions.ExportUnderlineFormatting
    //ExSummary:Shows how to export underline formatting as ++.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->set_Underline(Aspose::Words::Underline::Single);
    builder->Write(u"Lorem ipsum. Dolor sit amet.");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::MarkdownSaveOptions>();
    saveOptions->set_ExportUnderlineFormatting(true);
    doc->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.ExportUnderlineFormatting.md", saveOptions);
    //ExEnd:ExportUnderlineFormatting
}

namespace gtest_test
{

TEST_F(ExMarkdownSaveOptions, ExportUnderlineFormatting)
{
    s_instance->ExportUnderlineFormatting();
}

} // namespace gtest_test

void ExMarkdownSaveOptions::LinkExportMode()
{
    //ExStart:LinkExportMode
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:MarkdownSaveOptions.LinkExportMode
    //ExFor:MarkdownLinkExportMode
    //ExSummary:Shows how to links will be written to the .md file.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::Balloon, 100, 100);
    
    // Image will be written as reference:
    // ![ref1]
    // [ref1]: aw_ref.001.png
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::MarkdownSaveOptions>();
    saveOptions->set_LinkExportMode(Aspose::Words::Saving::MarkdownLinkExportMode::Reference);
    doc->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.LinkExportMode.Reference.md", saveOptions);
    
    // Image will be written as inline:
    // ![](aw_inline.001.png)
    saveOptions->set_LinkExportMode(Aspose::Words::Saving::MarkdownLinkExportMode::Inline);
    doc->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.LinkExportMode.Inline.md", saveOptions);
    //ExEnd:LinkExportMode
    
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"MarkdownSaveOptions.LinkExportMode.Inline.md");
    ASSERT_EQ(u"![](MarkdownSaveOptions.LinkExportMode.Inline.001.png)", outDocContents.Trim());
}

namespace gtest_test
{

TEST_F(ExMarkdownSaveOptions, LinkExportMode)
{
    s_instance->LinkExportMode();
}

} // namespace gtest_test

void ExMarkdownSaveOptions::ExportTableAsHtml()
{
    //ExStart:ExportTableAsHtml
    //GistId:bb594993b5fe48692541e16f4d354ac2
    //ExFor:MarkdownExportAsHtml
    //ExFor:MarkdownSaveOptions.ExportAsHtml
    //ExSummary:Shows how to export a table to Markdown as raw HTML.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Sample table:");
    
    // Create table.
    builder->InsertCell();
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Right);
    builder->Write(u"Cell1");
    builder->InsertCell();
    builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    builder->Write(u"Cell2");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::MarkdownSaveOptions>();
    saveOptions->set_ExportAsHtml(Aspose::Words::Saving::MarkdownExportAsHtml::Tables);
    
    doc->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.ExportTableAsHtml.md", saveOptions);
    //ExEnd:ExportTableAsHtml
    
    System::String newLine = System::Environment::get_NewLine();
    System::String outDocContents = System::IO::File::ReadAllText(get_ArtifactsDir() + u"MarkdownSaveOptions.ExportTableAsHtml.md");
    ASSERT_EQ(System::String::Format(u"Sample table:{0}<table cellspacing=\"0\" cellpadding=\"0\" style=\"width:100%; border:0.75pt solid #000000; border-collapse:collapse\">", newLine) + u"<tr><td style=\"border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">" + u"<p style=\"margin-top:0pt; margin-bottom:0pt; text-align:right; font-size:12pt\"><span style=\"font-family:'Times New Roman'\">Cell1</span></p>" + u"</td><td style=\"border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">" + u"<p style=\"margin-top:0pt; margin-bottom:0pt; text-align:center; font-size:12pt\"><span style=\"font-family:'Times New Roman'\">Cell2</span></p>" + u"</td></tr></table>", outDocContents.Trim());
}

namespace gtest_test
{

TEST_F(ExMarkdownSaveOptions, ExportTableAsHtml)
{
    s_instance->ExportTableAsHtml();
}

} // namespace gtest_test

void ExMarkdownSaveOptions::ImageResolution()
{
    //ExStart:ImageResolution
    //GistId:f86d49dc0e6781b93e576539a01e6ca2
    //ExFor:MarkdownSaveOptions.ImageResolution
    //ExSummary:Shows how to set the output resolution for images.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::MarkdownSaveOptions>();
    saveOptions->set_ImageResolution(300);
    
    doc->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.ImageResolution.md", saveOptions);
    //ExEnd:ImageResolution
}

namespace gtest_test
{

TEST_F(ExMarkdownSaveOptions, ImageResolution)
{
    s_instance->ImageResolution();
}

} // namespace gtest_test

void ExMarkdownSaveOptions::OfficeMathExportMode()
{
    //ExStart:OfficeMathExportMode
    //GistId:f86d49dc0e6781b93e576539a01e6ca2
    //ExFor:MarkdownSaveOptions.OfficeMathExportMode
    //ExFor:MarkdownOfficeMathExportMode
    //ExSummary:Shows how OfficeMath will be written to the document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::MarkdownSaveOptions>();
    saveOptions->set_OfficeMathExportMode(Aspose::Words::Saving::MarkdownOfficeMathExportMode::Image);
    
    doc->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.OfficeMathExportMode.md", saveOptions);
    //ExEnd:OfficeMathExportMode
}

namespace gtest_test
{

TEST_F(ExMarkdownSaveOptions, OfficeMathExportMode)
{
    s_instance->OfficeMathExportMode();
}

} // namespace gtest_test

void ExMarkdownSaveOptions::EmptyParagraphExportMode(Aspose::Words::Saving::MarkdownEmptyParagraphExportMode exportMode)
{
    //ExStart:EmptyParagraphExportMode
    //GistId:ad73e0dd58a8c2ae742bb64f8561df35
    //ExFor:MarkdownEmptyParagraphExportMode
    //ExFor:MarkdownSaveOptions.EmptyParagraphExportMode
    //ExSummary:Shows how to export empty paragraphs.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"First");
    builder->Writeln(u"\r\n\r\n\r\n");
    builder->Writeln(u"Last");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::MarkdownSaveOptions>();
    saveOptions->set_EmptyParagraphExportMode(exportMode);
    
    doc->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.EmptyParagraphExportMode.md", saveOptions);
    
    System::String result = System::IO::File::ReadAllText(get_ArtifactsDir() + u"MarkdownSaveOptions.EmptyParagraphExportMode.md");
    
    switch (exportMode)
    {
        case Aspose::Words::Saving::MarkdownEmptyParagraphExportMode::None:
            ASSERT_EQ(u"First\r\n\r\nLast\r\n", result);
            break;
        
        case Aspose::Words::Saving::MarkdownEmptyParagraphExportMode::EmptyLine:
            ASSERT_EQ(u"First\r\n\r\n\r\n\r\n\r\nLast\r\n\r\n", result);
            break;
        
        case Aspose::Words::Saving::MarkdownEmptyParagraphExportMode::MarkdownHardLineBreak:
            ASSERT_EQ(u"First\r\n\\\r\n\\\r\n\\\r\n\\\r\n\\\r\nLast\r\n<br>\r\n", result);
            break;
        
    }
    //ExEnd:EmptyParagraphExportMode
}

namespace gtest_test
{

using ExMarkdownSaveOptions_EmptyParagraphExportMode_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExMarkdownSaveOptions::EmptyParagraphExportMode)>::type;

struct ExMarkdownSaveOptions_EmptyParagraphExportMode : public ExMarkdownSaveOptions, public Aspose::Words::ApiExamples::ExMarkdownSaveOptions, public ::testing::WithParamInterface<ExMarkdownSaveOptions_EmptyParagraphExportMode_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::MarkdownEmptyParagraphExportMode::None),
            std::make_tuple(Aspose::Words::Saving::MarkdownEmptyParagraphExportMode::EmptyLine),
            std::make_tuple(Aspose::Words::Saving::MarkdownEmptyParagraphExportMode::MarkdownHardLineBreak),
        };
    }
};

TEST_P(ExMarkdownSaveOptions_EmptyParagraphExportMode, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->EmptyParagraphExportMode(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExMarkdownSaveOptions_EmptyParagraphExportMode, ::testing::ValuesIn(ExMarkdownSaveOptions_EmptyParagraphExportMode::TestCases()));

} // namespace gtest_test

void ExMarkdownSaveOptions::ExportOfficeMathAsLatex()
{
    //ExStart:ExportOfficeMathAsLatex
    //GistId:045648ef22da6b384ebcf0344717bfb5
    //ExFor:MarkdownSaveOptions.OfficeMathExportMode
    //ExFor:MarkdownOfficeMathExportMode
    //ExSummary:Shows how to export OfficeMath object as Latex.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::MarkdownSaveOptions>();
    saveOptions->set_OfficeMathExportMode(Aspose::Words::Saving::MarkdownOfficeMathExportMode::Latex);
    
    doc->Save(get_ArtifactsDir() + u"MarkdownSaveOptions.ExportOfficeMathAsLatex.md", saveOptions);
    //ExEnd:ExportOfficeMathAsLatex
    
    ASSERT_TRUE(Aspose::Words::ApiExamples::DocumentHelper::CompareDocs(get_ArtifactsDir() + u"MarkdownSaveOptions.ExportOfficeMathAsLatex.md", get_GoldsDir() + u"MarkdownSaveOptions.ExportOfficeMathAsLatex.Gold.md"));
}

namespace gtest_test
{

TEST_F(ExMarkdownSaveOptions, ExportOfficeMathAsLatex)
{
    s_instance->ExportOfficeMathAsLatex();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
