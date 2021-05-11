#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/BuildVersionInfo.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Fields/FieldToc.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Fields/FormField.h>
#include <Aspose.Words.Cpp/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Layout/LayoutCollector.h>
#include <Aspose.Words.Cpp/Lists/List.h>
#include <Aspose.Words.Cpp/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/PaperSize.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/CssStyleSheetType.h>
#include <Aspose.Words.Cpp/Saving/DocumentSplitCriteria.h>
#include <Aspose.Words.Cpp/Saving/ExportListLabels.h>
#include <Aspose.Words.Cpp/Saving/FontSavingArgs.h>
#include <Aspose.Words.Cpp/Saving/HtmlElementSizeOutputMode.h>
#include <Aspose.Words.Cpp/Saving/HtmlMetafileFormat.h>
#include <Aspose.Words.Cpp/Saving/HtmlOfficeMathOutputMode.h>
#include <Aspose.Words.Cpp/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/HtmlVersion.h>
#include <Aspose.Words.Cpp/Saving/IFontSavingCallback.h>
#include <Aspose.Words.Cpp/Saving/IImageSavingCallback.h>
#include <Aspose.Words.Cpp/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/PreferredWidth.h>
#include <drawing/image.h>
#include <drawing/rectangle_f.h>
#include <drawing/size.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/enum.h>
#include <system/enum_helpers.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/globalization/culture_info.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/file_info.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/io/path.h>
#include <system/io/search_option.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/math.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/encoding.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;

namespace ApiExamples {

class ExHtmlSaveOptions : public ApiExampleBase
{
public:
    void ExportPageMarginsEpub(SaveFormat saveFormat)
    {
        auto doc = MakeObject<Document>(MyDir + u"TextBoxes.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_SaveFormat(saveFormat);
        saveOptions->set_ExportPageMargins(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportPageMarginsEpub" + FileFormatUtil::SaveFormatToExtension(saveFormat), saveOptions);
    }

    void ExportOfficeMathEpub(SaveFormat saveFormat, HtmlOfficeMathOutputMode outputMode)
    {
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_OfficeMathOutputMode(outputMode);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportOfficeMathEpub" + FileFormatUtil::SaveFormatToExtension(saveFormat), saveOptions);
    }

    void ExportTextBoxAsSvgEpub(SaveFormat saveFormat, bool isTextBoxAsSvg)
    {
        ArrayPtr<String> dirFiles;

        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> textbox = builder->InsertShape(ShapeType::TextBox, 300, 100);
        builder->MoveTo(textbox->get_FirstParagraph());
        builder->Write(u"Hello world!");

        auto saveOptions = MakeObject<HtmlSaveOptions>(saveFormat);
        saveOptions->set_ExportTextBoxAsSvg(isTextBoxAsSvg);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportTextBoxAsSvgEpub" + FileFormatUtil::SaveFormatToExtension(saveFormat), saveOptions);

        switch (saveFormat)
        {
        case SaveFormat::Html:
            dirFiles =
                System::IO::Directory::GetFiles(ArtifactsDir, u"HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png", System::IO::SearchOption::AllDirectories);
            ASSERT_EQ(0, dirFiles->get_Length());
            return;

        case SaveFormat::Epub:
            dirFiles =
                System::IO::Directory::GetFiles(ArtifactsDir, u"HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png", System::IO::SearchOption::AllDirectories);
            ASSERT_EQ(0, dirFiles->get_Length());
            return;

        case SaveFormat::Mhtml:
            dirFiles =
                System::IO::Directory::GetFiles(ArtifactsDir, u"HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png", System::IO::SearchOption::AllDirectories);
            ASSERT_EQ(0, dirFiles->get_Length());
            return;

        default:
            break;
        }
    }

    void ControlListLabelsExport(ExportListLabels howExportListLabels)
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Aspose::Words::Lists::List> bulletedList = doc->get_Lists()->Add(ListTemplate::BulletDefault);
        builder->get_ListFormat()->set_List(bulletedList);
        builder->get_ParagraphFormat()->set_LeftIndent(72);
        builder->Writeln(u"Bulleted list item 1.");
        builder->Writeln(u"Bulleted list item 2.");
        builder->get_ParagraphFormat()->ClearFormatting();

        auto saveOptions = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        saveOptions->set_ExportListLabels(howExportListLabels);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ControlListLabelsExport.html", saveOptions);
    }

    void ExportUrlForLinkedImage(bool export_)
    {
        auto doc = MakeObject<Document>(MyDir + u"Linked image.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_ExportOriginalUrlForLinkedImages(export_);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

        ArrayPtr<String> dirFiles =
            System::IO::Directory::GetFiles(ArtifactsDir, u"HtmlSaveOptions.ExportUrlForLinkedImage.001.png", System::IO::SearchOption::AllDirectories);

        DocumentHelper::FindTextInFile(ArtifactsDir + u"HtmlSaveOptions.ExportUrlForLinkedImage.html",
                                       dirFiles->get_Length() == 0 ? String(u"<img src=\"http://www.aspose.com/images/aspose-logo.gif\"")
                                                                   : String(u"<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\""));
    }

    void ExportRoundtripInformation()
    {
        auto doc = MakeObject<Document>(MyDir + u"TextBoxes.docx");
        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_ExportRoundtripInformation(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.RoundtripInformation.html", saveOptions);
    }

    void RoundtripInformationDefaulValue()
    {
        auto saveOptions = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        ASPOSE_ASSERT_EQ(true, saveOptions->get_ExportRoundtripInformation());

        saveOptions = MakeObject<HtmlSaveOptions>(SaveFormat::Mhtml);
        ASPOSE_ASSERT_EQ(false, saveOptions->get_ExportRoundtripInformation());

        saveOptions = MakeObject<HtmlSaveOptions>(SaveFormat::Epub);
        ASPOSE_ASSERT_EQ(false, saveOptions->get_ExportRoundtripInformation());
    }

    void ExternalResourceSavingConfig()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_CssStyleSheetType(CssStyleSheetType::External);
        saveOptions->set_ExportFontResources(true);
        saveOptions->set_ResourceFolder(u"Resources");
        saveOptions->set_ResourceFolderAlias(u"https://www.aspose.com/");

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExternalResourceSavingConfig.html", saveOptions);

        ArrayPtr<String> imageFiles = System::IO::Directory::GetFiles(ArtifactsDir + u"Resources/", u"HtmlSaveOptions.ExternalResourceSavingConfig*.png",
                                                                      System::IO::SearchOption::AllDirectories);
        ASSERT_EQ(8, imageFiles->get_Length());

        ArrayPtr<String> fontFiles = System::IO::Directory::GetFiles(ArtifactsDir + u"Resources/", u"HtmlSaveOptions.ExternalResourceSavingConfig*.ttf",
                                                                     System::IO::SearchOption::AllDirectories);
        ASSERT_EQ(10, fontFiles->get_Length());

        ArrayPtr<String> cssFiles = System::IO::Directory::GetFiles(ArtifactsDir + u"Resources/", u"HtmlSaveOptions.ExternalResourceSavingConfig*.css",
                                                                    System::IO::SearchOption::AllDirectories);
        ASSERT_EQ(1, cssFiles->get_Length());

        DocumentHelper::FindTextInFile(ArtifactsDir + u"HtmlSaveOptions.ExternalResourceSavingConfig.html",
                                       u"<link href=\"https://www.aspose.com/HtmlSaveOptions.ExternalResourceSavingConfig.css\"");
    }

    void ConvertFontsAsBase64()
    {
        auto doc = MakeObject<Document>(MyDir + u"TextBoxes.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_CssStyleSheetType(CssStyleSheetType::External);
        saveOptions->set_ResourceFolder(u"Resources");
        saveOptions->set_ExportFontResources(true);
        saveOptions->set_ExportFontsAsBase64(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ConvertFontsAsBase64.html", saveOptions);
    }

    void Html5Support(HtmlVersion htmlVersion)
    {
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_HtmlVersion(htmlVersion);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.Html5Support.html", saveOptions);
    }

    void ExportFonts(bool exportAsBase64)
    {
        String fontsFolder = ArtifactsDir + u"HtmlSaveOptions.ExportFonts.Resources";

        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_ExportFontResources(true);
        saveOptions->set_FontsFolder(fontsFolder);
        saveOptions->set_ExportFontsAsBase64(exportAsBase64);

        switch (exportAsBase64)
        {
        case false:
            doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportFonts.False.html", saveOptions);
            ASSERT_FALSE(System::TestTools::IsEmpty(
                System::IO::Directory::GetFiles(fontsFolder, u"HtmlSaveOptions.ExportFonts.False.times.ttf", System::IO::SearchOption::AllDirectories)));
            System::IO::Directory::Delete(fontsFolder, true);
            break;

        case true:
            doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportFonts.True.html", saveOptions);
            ASSERT_FALSE(System::IO::Directory::Exists(fontsFolder));
            break;
        }
    }

    void ResourceFolderPriority()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_CssStyleSheetType(CssStyleSheetType::External);
        saveOptions->set_ExportFontResources(true);
        saveOptions->set_ResourceFolder(ArtifactsDir + u"Resources");
        saveOptions->set_ResourceFolderAlias(u"http://example.com/resources");

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ResourceFolderPriority.html", saveOptions);

        ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(ArtifactsDir + u"Resources", u"HtmlSaveOptions.ResourceFolderPriority.001.png",
                                                                                System::IO::SearchOption::AllDirectories)));
        ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(ArtifactsDir + u"Resources", u"HtmlSaveOptions.ResourceFolderPriority.002.png",
                                                                                System::IO::SearchOption::AllDirectories)));
        ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(
            ArtifactsDir + u"Resources", u"HtmlSaveOptions.ResourceFolderPriority.arial.ttf", System::IO::SearchOption::AllDirectories)));
        ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(ArtifactsDir + u"Resources", u"HtmlSaveOptions.ResourceFolderPriority.css",
                                                                                System::IO::SearchOption::AllDirectories)));
    }

    void ResourceFolderLowPriority()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_CssStyleSheetType(CssStyleSheetType::External);
        saveOptions->set_ExportFontResources(true);
        saveOptions->set_FontsFolder(ArtifactsDir + u"Fonts");
        saveOptions->set_ImagesFolder(ArtifactsDir + u"Images");
        saveOptions->set_ResourceFolder(ArtifactsDir + u"Resources");
        saveOptions->set_ResourceFolderAlias(u"http://example.com/resources");

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ResourceFolderLowPriority.html", saveOptions);

        ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(ArtifactsDir + u"Images", u"HtmlSaveOptions.ResourceFolderLowPriority.001.png",
                                                                                System::IO::SearchOption::AllDirectories)));
        ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(ArtifactsDir + u"Images", u"HtmlSaveOptions.ResourceFolderLowPriority.002.png",
                                                                                System::IO::SearchOption::AllDirectories)));
        ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(ArtifactsDir + u"Fonts", u"HtmlSaveOptions.ResourceFolderLowPriority.arial.ttf",
                                                                                System::IO::SearchOption::AllDirectories)));
        ASSERT_FALSE(System::TestTools::IsEmpty(System::IO::Directory::GetFiles(ArtifactsDir + u"Resources", u"HtmlSaveOptions.ResourceFolderLowPriority.css",
                                                                                System::IO::SearchOption::AllDirectories)));
    }

    void SvgMetafileFormat()
    {
        auto builder = MakeObject<DocumentBuilder>();

        builder->Write(u"Here is an SVG image: ");
        builder->InsertHtml(u"<svg height='210' width='500'>\r\n                    <polygon points='100,10 40,198 190,78 10,78 160,198' \r\n                  "
                            u"      style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n                  </svg> ");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_MetafileFormat(HtmlMetafileFormat::Png);
        builder->get_Document()->Save(ArtifactsDir + u"HtmlSaveOptions.SvgMetafileFormat.html", saveOptions);
    }

    void PngMetafileFormat()
    {
        auto builder = MakeObject<DocumentBuilder>();

        builder->Write(u"Here is an Png image: ");
        builder->InsertHtml(u"<svg height='210' width='500'>\r\n                    <polygon points='100,10 40,198 190,78 10,78 160,198' \r\n                  "
                            u"      style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n                  </svg> ");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_MetafileFormat(HtmlMetafileFormat::Png);
        builder->get_Document()->Save(ArtifactsDir + u"HtmlSaveOptions.PngMetafileFormat.html", saveOptions);
    }

    void EmfOrWmfMetafileFormat()
    {
        auto builder = MakeObject<DocumentBuilder>();

        builder->Write(u"Here is an image as is: ");
        builder->InsertHtml(
            u"<img src=\"data:image/png;base64,\r\n                    iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP\r\n                    "
            u"C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA\r\n                    "
            u"AAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J\r\n                    "
            u"REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq\r\n                    "
            u"ch9//q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0\r\n                    vr4MkhoXe0rZigAAAABJRU5ErkJggg==\" alt=\"Red dot\" />");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_MetafileFormat(HtmlMetafileFormat::EmfOrWmf);
        builder->get_Document()->Save(ArtifactsDir + u"HtmlSaveOptions.EmfOrWmfMetafileFormat.html", saveOptions);
    }

    void CssClassNamesPrefix()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.CssClassNamePrefix
        //ExSummary:Shows how to save a document to HTML, and add a prefix to all of its CSS class names.
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_CssStyleSheetType(CssStyleSheetType::External);
        saveOptions->set_CssClassNamePrefix(u"myprefix-");

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.CssClassNamePrefix.html");

        ASSERT_TRUE(outDocContents.Contains(u"<p class=\"myprefix-Header\">"));
        ASSERT_TRUE(outDocContents.Contains(u"<p class=\"myprefix-Footer\">"));

        outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.CssClassNamePrefix.css");

        ASSERT_TRUE(outDocContents.Contains(String(u".myprefix-Footer { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt }\r\n") +
                                            u".myprefix-Header { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt }\r\n"));
        //ExEnd
    }

    void CssClassNamesNotValidPrefix()
    {
        auto saveOptions = MakeObject<HtmlSaveOptions>();
        ASSERT_THROW(saveOptions->set_CssClassNamePrefix(u"@%-"), System::ArgumentException) << "The class name prefix must be a valid CSS identifier.";
    }

    void CssClassNamesNullPrefix()
    {
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_CssStyleSheetType(CssStyleSheetType::Embedded);
        saveOptions->set_CssClassNamePrefix(nullptr);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
    }

    void ContentIdScheme()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>(SaveFormat::Mhtml);
        saveOptions->set_PrettyFormat(true);
        saveOptions->set_ExportCidUrlsForMhtmlResources(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ContentIdScheme.mhtml", saveOptions);
    }

    void HeadingLevels()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.DocumentSplitHeadingLevel
        //ExSummary:Shows how to split an output HTML document by headings into several parts.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

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
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_DocumentSplitCriteria(DocumentSplitCriteria::HeadingParagraph);
        options->set_DocumentSplitHeadingLevel(2);

        // Our document has four headings of levels 1 - 2. One of those headings will not be
        // a split point since it is at the beginning of the document.
        // The saving operation will split our document at three places, into four smaller documents.
        doc->Save(ArtifactsDir + u"HtmlSaveOptions.HeadingLevels.html", options);

        doc = MakeObject<Document>(ArtifactsDir + u"HtmlSaveOptions.HeadingLevels.html");

        ASSERT_EQ(u"Heading #1", doc->GetText().Trim());

        doc = MakeObject<Document>(ArtifactsDir + u"HtmlSaveOptions.HeadingLevels-01.html");

        ASSERT_EQ(String(u"Heading #2\r") + u"Heading #3", doc->GetText().Trim());

        doc = MakeObject<Document>(ArtifactsDir + u"HtmlSaveOptions.HeadingLevels-02.html");

        ASSERT_EQ(u"Heading #4", doc->GetText().Trim());

        doc = MakeObject<Document>(ArtifactsDir + u"HtmlSaveOptions.HeadingLevels-03.html");

        ASSERT_EQ(String(u"Heading #5\r") + u"Heading #6", doc->GetText().Trim());
        //ExEnd
    }

    void NegativeIndent(bool allowNegativeIndent)
    {
        //ExStart
        //ExFor:HtmlElementSizeOutputMode
        //ExFor:HtmlSaveOptions.AllowNegativeIndent
        //ExFor:HtmlSaveOptions.TableWidthOutputMode
        //ExSummary:Shows how to preserve negative indents in the output .html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a table with a negative indent, which will push it to the left past the left page boundary.
        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 1");
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 2");
        builder->EndTable();
        table->set_LeftIndent(-36);
        table->set_PreferredWidth(PreferredWidth::FromPoints(144));

        builder->InsertBreak(BreakType::ParagraphBreak);

        // Insert a table with a positive indent, which will push the table to the right.
        table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 1");
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 2");
        builder->EndTable();
        table->set_LeftIndent(36);
        table->set_PreferredWidth(PreferredWidth::FromPoints(144));

        // When we save a document to HTML, Aspose.Words will only preserve negative indents
        // such as the one we have applied to the first table if we set the "AllowNegativeIndent" flag
        // in a SaveOptions object that we will pass to "true".
        auto options = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        options->set_AllowNegativeIndent(allowNegativeIndent);
        options->set_TableWidthOutputMode(HtmlElementSizeOutputMode::RelativeOnly);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.NegativeIndent.html", options);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.NegativeIndent.html");

        if (allowNegativeIndent)
        {
            ASSERT_TRUE(outDocContents.Contains(u"<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:-41.65pt; border:0.75pt solid #000000; "
                                                u"-aw-border:0.5pt single; border-collapse:collapse\">"));
            ASSERT_TRUE(outDocContents.Contains(u"<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; "
                                                u"-aw-border:0.5pt single; border-collapse:collapse\">"));
        }
        else
        {
            ASSERT_TRUE(outDocContents.Contains(
                u"<table cellspacing=\"0\" cellpadding=\"0\" style=\"border:0.75pt solid #000000; -aw-border:0.5pt single; border-collapse:collapse\">"));
            ASSERT_TRUE(outDocContents.Contains(u"<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; "
                                                u"-aw-border:0.5pt single; border-collapse:collapse\">"));
        }
        //ExEnd
    }

    void FolderAlias()
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
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<HtmlSaveOptions>();
        options->set_CssStyleSheetType(CssStyleSheetType::External);
        options->set_ExportFontResources(true);
        options->set_ImageResolution(72);
        options->set_FontResourcesSubsettingSizeThreshold(0);
        options->set_FontsFolder(ArtifactsDir + u"Fonts");
        options->set_ImagesFolder(ArtifactsDir + u"Images");
        options->set_ResourceFolder(ArtifactsDir + u"Resources");
        options->set_FontsFolderAlias(u"http://example.com/fonts");
        options->set_ImagesFolderAlias(u"http://example.com/images");
        options->set_ResourceFolderAlias(u"http://example.com/resources");
        options->set_ExportOriginalUrlForLinkedImages(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.FolderAlias.html", options);
        //ExEnd
    }

    //ExStart
    //ExFor:HtmlSaveOptions.ExportFontResources
    //ExFor:HtmlSaveOptions.FontSavingCallback
    //ExFor:IFontSavingCallback
    //ExFor:IFontSavingCallback.FontSaving
    //ExFor:FontSavingArgs
    //ExFor:FontSavingArgs.Bold
    //ExFor:FontSavingArgs.Document
    //ExFor:FontSavingArgs.FontFamilyName
    //ExFor:FontSavingArgs.FontFileName
    //ExFor:FontSavingArgs.FontStream
    //ExFor:FontSavingArgs.IsExportNeeded
    //ExFor:FontSavingArgs.IsSubsettingNeeded
    //ExFor:FontSavingArgs.Italic
    //ExFor:FontSavingArgs.KeepFontStreamOpen
    //ExFor:FontSavingArgs.OriginalFileName
    //ExFor:FontSavingArgs.OriginalFileSize
    //ExSummary:Shows how to define custom logic for exporting fonts when saving to HTML.
    void SaveExportedFonts()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Configure a SaveOptions object to export fonts to separate files.
        // Set a callback that will handle font saving in a custom manner.
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportFontResources(true);
        options->set_FontSavingCallback(MakeObject<ExHtmlSaveOptions::HandleFontSaving>());

        // The callback will export .ttf files and save them alongside the output document.
        doc->Save(ArtifactsDir + u"HtmlSaveOptions.SaveExportedFonts.html", options);

        std::function<bool(String s)> endsWithTtf = [](String s)
        {
            return s.EndsWith(u".ttf");
        };

        for (String fontFilename : System::Array<String>::FindAll(System::IO::Directory::GetFiles(ArtifactsDir), endsWithTtf))
        {
            std::cout << fontFilename << std::endl;
        }

        ASSERT_EQ(10, System::Array<String>::FindAll(System::IO::Directory::GetFiles(ArtifactsDir), endsWithTtf)->get_Length());
        //ExSkip
    }

    /// <summary>
    /// Prints information about exported fonts and saves them in the same local system folder as their output .html.
    /// </summary>
    class HandleFontSaving : public IFontSavingCallback
    {
    public:
        void FontSaving(SharedPtr<FontSavingArgs> args) override
        {
            std::cout << "Font:\t" << args->get_FontFamilyName();
            if (args->get_Bold())
            {
                std::cout << ", bold";
            }
            if (args->get_Italic())
            {
                std::cout << ", italic";
            }
            std::cout << "\nSource:\t" << args->get_OriginalFileName() << ", " << args->get_OriginalFileSize() << " bytes\n" << std::endl;

            // We can also access the source document from here.
            ASSERT_TRUE(args->get_Document()->get_OriginalFileName().EndsWith(u"Rendering.docx"));

            ASSERT_TRUE(args->get_IsExportNeeded());
            ASSERT_TRUE(args->get_IsSubsettingNeeded());

            // There are two ways of saving an exported font.
            // 1 -  Save it to a local file system location:
            args->set_FontFileName(args->get_OriginalFileName().Split(MakeArray<char16_t>({System::IO::Path::DirectorySeparatorChar}))->LINQ_Last());

            // 2 -  Save it to a stream:
            args->set_FontStream(MakeObject<System::IO::FileStream>(
                ArtifactsDir + args->get_OriginalFileName().Split(MakeArray<char16_t>({System::IO::Path::DirectorySeparatorChar}))->LINQ_Last(),
                System::IO::FileMode::Create));
            ASSERT_FALSE(args->get_KeepFontStreamOpen());
        }
    };
    //ExEnd

    void HtmlVersions(HtmlVersion htmlVersion)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.#ctor(SaveFormat)
        //ExFor:HtmlSaveOptions.HtmlVersion
        //ExFor:HtmlVersion
        //ExSummary:Shows how to save a document to a specific version of HTML.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        options->set_HtmlVersion(htmlVersion);
        options->set_PrettyFormat(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.HtmlVersions.html", options);

        // Our HTML documents will have minor differences to be compatible with different HTML versions.
        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.HtmlVersions.html");

        switch (htmlVersion)
        {
        case HtmlVersion::Html5:
            ASSERT_TRUE(outDocContents.Contains(u"<a id=\"_Toc76372689\"></a>"));
            ASSERT_TRUE(outDocContents.Contains(u"<a id=\"_Toc76372689\"></a>"));
            ASSERT_TRUE(outDocContents.Contains(u"<table style=\"border-collapse:collapse\">"));
            break;

        case HtmlVersion::Xhtml:
            ASSERT_TRUE(outDocContents.Contains(u"<a name=\"_Toc76372689\"></a>"));
            ASSERT_TRUE(outDocContents.Contains(u"<ul type=\"disc\" style=\"margin:0pt; padding-left:0pt\">"));
            ASSERT_TRUE(outDocContents.Contains(u"<table cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse\">"));
            break;
        }
        //ExEnd
    }

    void ExportXhtmlTransitional(bool showDoctypeDeclaration)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportXhtmlTransitional
        //ExFor:HtmlSaveOptions.HtmlVersion
        //ExFor:HtmlVersion
        //ExSummary:Shows how to display a DOCTYPE heading when converting documents to the Xhtml 1.0 transitional standard.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");

        auto options = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        options->set_HtmlVersion(HtmlVersion::Xhtml);
        options->set_ExportXhtmlTransitional(showDoctypeDeclaration);
        options->set_PrettyFormat(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportXhtmlTransitional.html", options);

        // Our document will only contain a DOCTYPE declaration heading if we have set the "ExportXhtmlTransitional" flag to "true".
        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.ExportXhtmlTransitional.html");

        if (showDoctypeDeclaration)
        {
            ASSERT_TRUE(outDocContents.Contains(String(u"<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"no\"?>\r\n") +
                                                u"<!DOCTYPE html\r\nPUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\"\r\n       "
                                                u"\"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">\r\n" +
                                                u"<html xmlns=\"http://www.w3.org/1999/xhtml\">"));
        }
        else
        {
            ASSERT_TRUE(outDocContents.Contains(u"<html>"));
        }
        //ExEnd
    }

    void EpubHeadings()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.EpubNavigationMapLevel
        //ExSummary:Shows how to filter headings that appear in the navigation panel of a saved Epub document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

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

        // Epub readers typically create a table of contents for their documents.
        // Each paragraph with a "Heading" style in the document will create an entry in this table of contents.
        // We can use the "EpubNavigationMapLevel" property to set a maximum heading level.
        // The Epub reader will not add headings with a level above the one we specify to the contents table.
        auto options = MakeObject<HtmlSaveOptions>(SaveFormat::Epub);
        options->set_EpubNavigationMapLevel(2);

        // Our document has six headings, two of which are above level 2.
        // The table of contents for this document will have four entries.
        doc->Save(ArtifactsDir + u"HtmlSaveOptions.EpubHeadings.epub", options);
        //ExEnd
    }

    void Doc2EpubSaveOptions()
    {
        //ExStart
        //ExFor:DocumentSplitCriteria
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.#ctor()
        //ExFor:HtmlSaveOptions.Encoding
        //ExFor:HtmlSaveOptions.DocumentSplitCriteria
        //ExFor:HtmlSaveOptions.ExportDocumentProperties
        //ExFor:HtmlSaveOptions.SaveFormat
        //ExFor:SaveOptions
        //ExFor:SaveOptions.SaveFormat
        //ExSummary:Shows how to use a specific encoding when saving a document to .epub.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Use a SaveOptions object to specify the encoding for a document that we will save.
        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_SaveFormat(SaveFormat::Epub);
        saveOptions->set_Encoding(System::Text::Encoding::get_UTF8());

        // By default, an output .epub document will have all its contents in one HTML part.
        // A split criterion allows us to segment the document into several HTML parts.
        // We will set the criteria to split the document into heading paragraphs.
        // This is useful for readers who cannot read HTML files more significant than a specific size.
        saveOptions->set_DocumentSplitCriteria(DocumentSplitCriteria::HeadingParagraph);

        // Specify that we want to export document properties.
        saveOptions->set_ExportDocumentProperties(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.Doc2EpubSaveOptions.epub", saveOptions);
        //ExEnd
    }

    void ContentIdUrls(bool exportCidUrlsForMhtmlResources)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportCidUrlsForMhtmlResources
        //ExSummary:Shows how to enable content IDs for output MHTML documents.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Setting this flag will replace "Content-Location" tags
        // with "Content-ID" tags for each resource from the input document.
        auto options = MakeObject<HtmlSaveOptions>(SaveFormat::Mhtml);
        options->set_ExportCidUrlsForMhtmlResources(exportCidUrlsForMhtmlResources);
        options->set_CssStyleSheetType(CssStyleSheetType::External);
        options->set_ExportFontResources(true);
        options->set_PrettyFormat(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ContentIdUrls.mht", options);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.ContentIdUrls.mht");

        if (exportCidUrlsForMhtmlResources)
        {
            ASSERT_TRUE(outDocContents.Contains(u"Content-ID: <document.html>"));
            ASSERT_TRUE(outDocContents.Contains(u"<link href=3D\"cid:styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />"));
            ASSERT_TRUE(outDocContents.Contains(u"@font-face { font-family:'Arial Black'; src:url('cid:ariblk.ttf') }"));
            ASSERT_TRUE(outDocContents.Contains(u"<img src=3D\"cid:image.003.jpeg\" width=3D\"351\" height=3D\"180\" alt=3D\"\" />"));
        }
        else
        {
            ASSERT_TRUE(outDocContents.Contains(u"Content-Location: document.html"));
            ASSERT_TRUE(outDocContents.Contains(u"<link href=3D\"styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />"));
            ASSERT_TRUE(outDocContents.Contains(u"@font-face { font-family:'Arial Black'; src:url('ariblk.ttf') }"));
            ASSERT_TRUE(outDocContents.Contains(u"<img src=3D\"image.003.jpeg\" width=3D\"351\" height=3D\"180\" alt=3D\"\" />"));
        }
        //ExEnd
    }

    void DropDownFormField(bool exportDropDownFormFieldAsText)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportDropDownFormFieldAsText
        //ExSummary:Shows how to get drop-down combo box form fields to blend in with paragraph text when saving to html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to insert a combo box with the value "Two" selected.
        builder->InsertComboBox(u"MyComboBox", MakeArray<String>({u"One", u"Two", u"Three"}), 1);

        // The "ExportDropDownFormFieldAsText" flag of this SaveOptions object allows us to
        // control how saving the document to HTML treats drop-down combo boxes.
        // Setting it to "true" will convert each combo box into simple text
        // that displays the combo box's currently selected value, effectively freezing it.
        // Setting it to "false" will preserve the functionality of the combo box using <select> and <option> tags.
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportDropDownFormFieldAsText(exportDropDownFormFieldAsText);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.DropDownFormField.html", options);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.DropDownFormField.html");

        if (exportDropDownFormFieldAsText)
        {
            ASSERT_TRUE(outDocContents.Contains(u"<span>Two</span>"));
        }
        else
        {
            ASSERT_TRUE(outDocContents.Contains(String(u"<select name=\"MyComboBox\">") + u"<option>One</option>" +
                                                u"<option selected=\"selected\">Two</option>" + u"<option>Three</option>" + u"</select>"));
        }
        //ExEnd
    }

    void ExportImagesAsBase64(bool exportItemsAsBase64)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportFontsAsBase64
        //ExFor:HtmlSaveOptions.ExportImagesAsBase64
        //ExSummary:Shows how to save a .html document with images embedded inside it.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportImagesAsBase64(exportItemsAsBase64);
        options->set_PrettyFormat(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportImagesAsBase64.html", options);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.ExportImagesAsBase64.html");

        ASSERT_TRUE(exportItemsAsBase64 ? outDocContents.Contains(u"<img src=\"data:image/png;base64")
                                        : outDocContents.Contains(u"<img src=\"HtmlSaveOptions.ExportImagesAsBase64.001.png\""));
        //ExEnd
    }

    void ExportFontsAsBase64()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportFontsAsBase64
        //ExFor:HtmlSaveOptions.ExportImagesAsBase64
        //ExSummary:Shows how to embed fonts inside a saved HTML document.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportFontsAsBase64(true);
        options->set_CssStyleSheetType(CssStyleSheetType::Embedded);
        options->set_PrettyFormat(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportFontsAsBase64.html", options);
        //ExEnd
    }

    void ExportLanguageInformation(bool exportLanguageInformation)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportLanguageInformation
        //ExSummary:Shows how to preserve language information when saving to .html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use the builder to write text while formatting it in different locales.
        builder->get_Font()->set_LocaleId(MakeObject<System::Globalization::CultureInfo>(u"en-US")->get_LCID());
        builder->Writeln(u"Hello world!");

        builder->get_Font()->set_LocaleId(MakeObject<System::Globalization::CultureInfo>(u"en-GB")->get_LCID());
        builder->Writeln(u"Hello again!");

        builder->get_Font()->set_LocaleId(MakeObject<System::Globalization::CultureInfo>(u"ru-RU")->get_LCID());
        builder->Write(u"Привет, мир!");

        // When saving the document to HTML, we can pass a SaveOptions object
        // to either preserve or discard each formatted text's locale.
        // If we set the "ExportLanguageInformation" flag to "true",
        // the output HTML document will contain the locales in "lang" attributes of <span> tags.
        // If we set the "ExportLanguageInformation" flag to "false',
        // the text in the output HTML document will not contain any locale information.
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportLanguageInformation(exportLanguageInformation);
        options->set_PrettyFormat(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportLanguageInformation.html", options);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.ExportLanguageInformation.html");

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

    void List(ExportListLabels exportListLabels)
    {
        //ExStart
        //ExFor:ExportListLabels
        //ExFor:HtmlSaveOptions.ExportListLabels
        //ExSummary:Shows how to configure list exporting to HTML.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(ListTemplate::NumberDefault);
        builder->get_ListFormat()->set_List(list);

        builder->Writeln(u"Default numbered list item 1.");
        builder->Writeln(u"Default numbered list item 2.");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"Default numbered list item 3.");
        builder->get_ListFormat()->RemoveNumbers();

        list = doc->get_Lists()->Add(ListTemplate::OutlineHeadingsLegal);
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
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportListLabels(exportListLabels);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.List.html", options);
        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.List.html");

        switch (exportListLabels)
        {
        case ExportListLabels::AsInlineText:
            ASSERT_TRUE(outDocContents.Contains(
                String(u"<p style=\"margin-top:0pt; margin-left:72pt; margin-bottom:0pt; text-indent:-18pt; -aw-import:list-item; -aw-list-level-number:1; "
                       u"-aw-list-number-format:'%1.'; -aw-list-number-styles:'lowerLetter'; -aw-list-number-values:'1'; -aw-list-padding-sml:9.67pt\">") +
                u"<span style=\"-aw-import:ignore\">" + u"<span>a.</span>" +
                u"<span style=\"width:9.67pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; "
                u"</span>" +
                u"</span>" + u"<span>Default numbered list item 3.</span>" + u"</p>"));
            ASSERT_TRUE(
                outDocContents.Contains(String(u"<p style=\"margin-top:0pt; margin-left:43.2pt; margin-bottom:0pt; text-indent:-43.2pt; -aw-import:list-item; "
                                               u"-aw-list-level-number:3; -aw-list-number-format:'%0.%1.%2.%3'; -aw-list-number-styles:'decimal decimal "
                                               u"decimal decimal'; -aw-list-number-values:'2 1 1 1'; -aw-list-padding-sml:10.2pt\">") +
                                        u"<span style=\"-aw-import:ignore\">" + u"<span>2.1.1.1</span>" +
                                        u"<span style=\"width:10.2pt; font:7pt 'Times New Roman'; display:inline-block; "
                                        u"-aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                                        u"</span>" + u"<span>Outline legal heading list item 5.</span>" + u"</p>"));
            break;

        case ExportListLabels::Auto:
            ASSERT_TRUE(outDocContents.Contains(String(u"<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">") +
                                                u"<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" + u"<span>Default numbered list item 3.</span>" +
                                                u"</li>" + u"</ol>"));
            ASSERT_TRUE(outDocContents.Contains(
                String(
                    u"<p style=\"margin-top:0pt; margin-left:43.2pt; margin-bottom:0pt; text-indent:-43.2pt; -aw-import:list-item; -aw-list-level-number:3; ") +
                u"-aw-list-number-format:'%0.%1.%2.%3'; -aw-list-number-styles:'decimal decimal decimal decimal'; " +
                u"-aw-list-number-values:'2 1 1 1'; -aw-list-padding-sml:10.2pt\">" + u"<span style=\"-aw-import:ignore\">" + u"<span>2.1.1.1</span>" +
                u"<span style=\"width:10.2pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; "
                u"</span>" +
                u"</span>" + u"<span>Outline legal heading list item 5.</span>" + u"</p>"));
            break;

        case ExportListLabels::ByHtmlTags:
            ASSERT_TRUE(outDocContents.Contains(String(u"<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">") +
                                                u"<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" + u"<span>Default numbered list item 3.</span>" +
                                                u"</li>" + u"</ol>"));
            ASSERT_TRUE(outDocContents.Contains(String(u"<ol type=\"1\" class=\"awlist3\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">") +
                                                u"<li style=\"margin-left:7.2pt; text-indent:-43.2pt; -aw-list-padding-sml:10.2pt\">" +
                                                u"<span style=\"width:10.2pt; font:7pt 'Times New Roman'; display:inline-block; "
                                                u"-aw-import:ignore\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                                                u"<span>Outline legal heading list item 5.</span>" + u"</li>" + u"</ol>"));
            break;
        }
        //ExEnd
    }

    void ExportPageMargins(bool exportPageMargins)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportPageMargins
        //ExSummary:Shows how to show out-of-bounds objects in output HTML documents.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a builder to insert a shape with no wrapping.
        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Cube, 200, 200);

        shape->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);
        shape->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);
        shape->set_WrapType(WrapType::None);

        // Negative shape position values may place the shape outside of page boundaries.
        // If we export this to HTML, the shape will appear truncated.
        shape->set_Left(-150);

        // When saving the document to HTML, we can pass a SaveOptions object
        // to decide whether to adjust the page to display out-of-bounds objects fully.
        // If we set the "ExportPageMargins" flag to "true", the shape will be fully visible in the output HTML.
        // If we set the "ExportPageMargins" flag to "false",
        // our document will display the shape truncated as we would see it in Microsoft Word.
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportPageMargins(exportPageMargins);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportPageMargins.html", options);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.ExportPageMargins.html");

        if (exportPageMargins)
        {
            ASSERT_TRUE(outDocContents.Contains(u"<style type=\"text/css\">div.Section1 { margin:70.85pt }</style>"));
            ASSERT_TRUE(outDocContents.Contains(u"<div class=\"Section1\"><p style=\"margin-top:0pt; margin-left:151pt; margin-bottom:0pt\">"));
        }
        else
        {
            ASSERT_FALSE(outDocContents.Contains(u"style type=\"text/css\">"));
            ASSERT_TRUE(outDocContents.Contains(u"<div><p style=\"margin-top:0pt; margin-left:221.85pt; margin-bottom:0pt\">"));
        }
        //ExEnd
    }

    void ExportPageSetup(bool exportPageSetup)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportPageSetup
        //ExSummary:Shows how decide whether to preserve section structure/page setup information when saving to HTML.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Section 1");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Write(u"Section 2");

        SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
        pageSetup->set_TopMargin(36.0);
        pageSetup->set_BottomMargin(36.0);
        pageSetup->set_PaperSize(PaperSize::A5);

        // When saving the document to HTML, we can pass a SaveOptions object
        // to decide whether to preserve or discard page setup settings.
        // If we set the "ExportPageSetup" flag to "true", the output HTML document will contain our page setup configuration.
        // If we set the "ExportPageSetup" flag to "false", the save operation will discard our page setup settings
        // for the first section, and both sections will look identical.
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportPageSetup(exportPageSetup);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportPageSetup.html", options);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.ExportPageSetup.html");

        if (exportPageSetup)
        {
            ASSERT_TRUE(outDocContents.Contains(String(u"<style type=\"text/css\">") + u"@page Section1 { size:419.55pt 595.3pt; margin:36pt 70.85pt }" +
                                                u"@page Section2 { size:612pt 792pt; margin:70.85pt }" +
                                                u"div.Section1 { page:Section1 }div.Section2 { page:Section2 }" + u"</style>"));

            ASSERT_TRUE(outDocContents.Contains(String(u"<div class=\"Section1\">") + u"<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                                                u"<span>Section 1</span>" + u"</p>" + u"</div>"));
        }
        else
        {
            ASSERT_FALSE(outDocContents.Contains(u"style type=\"text/css\">"));

            ASSERT_TRUE(outDocContents.Contains(String(u"<div>") + u"<p style=\"margin-top:0pt; margin-bottom:0pt\">" + u"<span>Section 1</span>" + u"</p>" +
                                                u"</div>"));
        }
        //ExEnd
    }

    void RelativeFontSize(bool exportRelativeFontSize)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportRelativeFontSize
        //ExSummary:Shows how to use relative font sizes when saving to .html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

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
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportRelativeFontSize(exportRelativeFontSize);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.RelativeFontSize.html", options);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.RelativeFontSize.html");

        if (exportRelativeFontSize)
        {
            ASSERT_TRUE(outDocContents.Contains(String(u"<body style=\"font-family:'Times New Roman'\">") + u"<div>" +
                                                u"<p style=\"margin-top:0pt; margin-bottom:0pt\">" + u"<span>Default font size, </span>" + u"</p>" +
                                                u"<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:2em\">" + u"<span>2x default font size,</span>" +
                                                u"</p>" + u"<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:8em\">" +
                                                u"<span>8x default font size</span>" + u"</p>" + u"</div>" + u"</body>"));
        }
        else
        {
            ASSERT_TRUE(outDocContents.Contains(String(u"<body style=\"font-family:'Times New Roman'; font-size:12pt\">") + u"<div>" +
                                                u"<p style=\"margin-top:0pt; margin-bottom:0pt\">" + u"<span>Default font size, </span>" + u"</p>" +
                                                u"<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:24pt\">" + u"<span>2x default font size,</span>" +
                                                u"</p>" + u"<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:96pt\">" +
                                                u"<span>8x default font size</span>" + u"</p>" + u"</div>" + u"</body>"));
        }
        //ExEnd
    }

    void ExportTextBox(bool exportTextBoxAsSvg)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportTextBoxAsSvg
        //ExSummary:Shows how to export text boxes as scalable vector graphics.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> textBox = builder->InsertShape(ShapeType::TextBox, 100.0, 60.0);
        builder->MoveTo(textBox->get_FirstParagraph());
        builder->Write(u"My text box");

        // When we save the document to HTML, we can pass a SaveOptions object
        // to determine how the saving operation will export text box shapes.
        // If we set the "ExportTextBoxAsSvg" flag to "true",
        // the save operation will convert shapes with text into SVG objects.
        // If we set the "ExportTextBoxAsSvg" flag to "false",
        // the save operation will convert shapes with text into images.
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportTextBoxAsSvg(exportTextBoxAsSvg);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportTextBox.html", options);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.ExportTextBox.html");

        if (exportTextBoxAsSvg)
        {
            ASSERT_TRUE(outDocContents.Contains(
                String(u"<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">") +
                u"<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"133\" height=\"80\">"));
        }
        else
        {
            ASSERT_TRUE(outDocContents.Contains(
                String(u"<p style=\"margin-top:0pt; margin-bottom:0pt\">") +
                u"<img src=\"HtmlSaveOptions.ExportTextBox.001.png\" width=\"136\" height=\"83\" alt=\"\" " +
                u"style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" + u"</p>"));
        }
        //ExEnd
    }

    void RoundTripInformation(bool exportRoundtripInformation)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportRoundtripInformation
        //ExSummary:Shows how to preserve hidden elements when converting to .html.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

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
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportRoundtripInformation(exportRoundtripInformation);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.RoundTripInformation.html", options);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.RoundTripInformation.html");
        doc = MakeObject<Document>(ArtifactsDir + u"HtmlSaveOptions.RoundTripInformation.html");

        if (exportRoundtripInformation)
        {
            ASSERT_TRUE(outDocContents.Contains(u"<div style=\"-aw-headerfooter-type:header-primary; clear:both\">"));
            ASSERT_TRUE(outDocContents.Contains(u"<span style=\"-aw-import:ignore\">&#xa0;</span>"));

            ASSERT_TRUE(outDocContents.Contains(String(u"td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; ") +
                                                u"padding-right:2.4pt; padding-left:5.03pt; vertical-align:top; " +
                                                u"-aw-border-bottom:0.5pt single; -aw-border-left:0.5pt single; -aw-border-top:0.5pt single\">"));

            ASSERT_TRUE(outDocContents.Contains(
                u"<li style=\"margin-left:30.2pt; padding-left:5.8pt; -aw-font-family:'Courier New'; -aw-font-weight:normal; -aw-number-format:'o'\">"));

            ASSERT_TRUE(
                outDocContents.Contains(String(u"<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"351\" height=\"180\" alt=\"\" ") +
                                        u"style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />"));

            ASSERT_TRUE(outDocContents.Contains(String(u"<span>Page number </span>") + u"<span style=\"-aw-field-start:true\"></span>" +
                                                u"<span style=\"-aw-field-code:' PAGE   \\\\* MERGEFORMAT '\"></span>" +
                                                u"<span style=\"-aw-field-separator:true\"></span>" + u"<span>1</span>" +
                                                u"<span style=\"-aw-field-end:true\"></span>"));

            ASSERT_EQ(1, doc->get_Range()->get_Fields()->LINQ_Count([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldPage; }));
        }
        else
        {
            ASSERT_TRUE(outDocContents.Contains(u"<div style=\"clear:both\">"));
            ASSERT_TRUE(outDocContents.Contains(u"<span>&#xa0;</span>"));

            ASSERT_TRUE(
                outDocContents.Contains(String(u"<td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; ") +
                                        u"padding-right:2.4pt; padding-left:5.03pt; vertical-align:top\">"));

            ASSERT_TRUE(outDocContents.Contains(u"<li style=\"margin-left:30.2pt; padding-left:5.8pt\">"));

            ASSERT_TRUE(outDocContents.Contains(u"<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"351\" height=\"180\" alt=\"\" />"));

            ASSERT_TRUE(outDocContents.Contains(u"<span>Page number 1</span>"));

            ASSERT_EQ(0, doc->get_Range()->get_Fields()->LINQ_Count([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldPage; }));
        }
        //ExEnd
    }

    void ExportTocPageNumbers(bool exportTocPageNumbers)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportTocPageNumbers
        //ExSummary:Shows how to display page numbers when saving a document with a table of contents to .html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a table of contents, and then populate the document with paragraphs formatted using a "Heading"
        // style that the table of contents will pick up as entries. Each entry will display the heading paragraph on the left,
        // and the page number that contains the heading on the right.
        auto fieldToc = System::DynamicCast<FieldToc>(builder->InsertField(FieldType::FieldTOC, true));

        builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(u"Heading 1"));
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Entry 1");
        builder->Writeln(u"Entry 2");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Entry 3");
        builder->InsertBreak(BreakType::PageBreak);
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
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportTocPageNumbers(exportTocPageNumbers);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportTocPageNumbers.html", options);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.ExportTocPageNumbers.html");

        if (exportTocPageNumbers)
        {
            ASSERT_TRUE(outDocContents.Contains(
                String(u"<p style=\"margin-top:0pt; margin-bottom:0pt\">") + u"<span>Entry 1</span>" +
                u"<span style=\"width:428.14pt; font-family:'Lucida Console'; font-size:10pt; display:inline-block; -aw-font-family:'Times New Roman'; " +
                u"-aw-tabstop-align:right; -aw-tabstop-leader:dots; "
                u"-aw-tabstop-pos:469.8pt\">.......................................................................</span>" +
                u"<span>2</span>" + u"</p>"));
        }
        else
        {
            ASSERT_TRUE(outDocContents.Contains(String(u"<p style=\"margin-top:0pt; margin-bottom:0pt\">") + u"<span>Entry 1</span>" + u"</p>"));
        }
        //ExEnd
    }

    void FontSubsetting(int fontResourcesSubsettingSizeThreshold)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.FontResourcesSubsettingSizeThreshold
        //ExSummary:Shows how to work with font subsetting.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

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
        String fontsFolder = ArtifactsDir + u"HtmlSaveOptions.FontSubsetting.Fonts";

        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportFontResources(true);
        options->set_FontsFolder(fontsFolder);
        options->set_FontResourcesSubsettingSizeThreshold(fontResourcesSubsettingSizeThreshold);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.FontSubsetting.html", options);

        ArrayPtr<String> fontFileNames = System::IO::Directory::GetFiles(fontsFolder)->LINQ_Where([](String s) { return s.EndsWith(u".ttf"); })->LINQ_ToArray();

        ASSERT_EQ(3, fontFileNames->get_Length());

        for (String filename : fontFileNames)
        {
            // By default, the .ttf files for each of our three fonts will be over 700MB.
            // Subsetting will reduce them all to under 30MB.
            auto fontFileInfo = MakeObject<System::IO::FileInfo>(filename);

            ASSERT_TRUE(fontFileInfo->get_Length() > 700000 || fontFileInfo->get_Length() < 30000);
            ASSERT_TRUE(System::Math::Max(fontResourcesSubsettingSizeThreshold, 30000) > MakeObject<System::IO::FileInfo>(filename)->get_Length());
        }

        //ExEnd
    }

    void MetafileFormat(HtmlMetafileFormat htmlMetafileFormat)
    {
        //ExStart
        //ExFor:HtmlMetafileFormat
        //ExFor:HtmlSaveOptions.MetafileFormat
        //ExSummary:Shows how to convert SVG objects to a different format when saving HTML documents.
        String html =
            u"<html>\r\n                    <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>\r\n                        "
            u"<text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>\r\n                    </svg>\r\n                </html>";

        auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(html)));

        // This document contains a <svg> element in the form of text.
        // When we save the document to HTML, we can pass a SaveOptions object
        // to determine how the saving operation handles this object.
        // Setting the "MetafileFormat" property to "HtmlMetafileFormat.Png" to convert it to a PNG image.
        // Setting the "MetafileFormat" property to "HtmlMetafileFormat.Svg" preserve it as a SVG object.
        // Setting the "MetafileFormat" property to "HtmlMetafileFormat.EmfOrWmf" to convert it to a metafile.
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_MetafileFormat(htmlMetafileFormat);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.MetafileFormat.html", options);

        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.MetafileFormat.html");

        switch (htmlMetafileFormat)
        {
        case HtmlMetafileFormat::Png:
            ASSERT_TRUE(outDocContents.Contains(
                String(u"<p style=\"margin-top:0pt; margin-bottom:0pt\">") +
                u"<img src=\"HtmlSaveOptions.MetafileFormat.001.png\" width=\"500\" height=\"40\" alt=\"\" " +
                u"style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" + u"</p>"));
            break;

        case HtmlMetafileFormat::Svg:
            ASSERT_TRUE(outDocContents.Contains(
                String(u"<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">") +
                u"<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"499\" height=\"40\">"));
            break;

        case HtmlMetafileFormat::EmfOrWmf:
            ASSERT_TRUE(outDocContents.Contains(
                String(u"<p style=\"margin-top:0pt; margin-bottom:0pt\">") +
                u"<img src=\"HtmlSaveOptions.MetafileFormat.001.emf\" width=\"500\" height=\"40\" alt=\"\" " +
                u"style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" + u"</p>"));
            break;
        }
        //ExEnd
    }

    void OfficeMathOutputMode(HtmlOfficeMathOutputMode htmlOfficeMathOutputMode)
    {
        //ExStart
        //ExFor:HtmlOfficeMathOutputMode
        //ExFor:HtmlSaveOptions.OfficeMathOutputMode
        //ExSummary:Shows how to specify how to export Microsoft OfficeMath objects to HTML.
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

        // When we save the document to HTML, we can pass a SaveOptions object
        // to determine how the saving operation handles OfficeMath objects.
        // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.Image"
        // will render each OfficeMath object into an image.
        // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.MathML"
        // will convert each OfficeMath object into MathML.
        // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.Text"
        // will represent each OfficeMath formula using plain HTML text.
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_OfficeMathOutputMode(htmlOfficeMathOutputMode);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.OfficeMathOutputMode.html", options);
        String outDocContents = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.OfficeMathOutputMode.html");

        switch (htmlOfficeMathOutputMode)
        {
        case HtmlOfficeMathOutputMode::Image:
            ASSERT_TRUE(
                System::Text::RegularExpressions::Regex::Match(
                    outDocContents,
                    String(u"<p style=\"margin-top:0pt; margin-bottom:10pt\">") +
                        u"<img src=\"HtmlSaveOptions.OfficeMathOutputMode.001.png\" width=\"159\" height=\"19\" alt=\"\" style=\"vertical-align:middle; " +
                        u"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" + u"</p>")
                    ->get_Success());
            break;

        case HtmlOfficeMathOutputMode::MathML:
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(outDocContents, String(u"<p style=\"margin-top:0pt; margin-bottom:10pt\">") +
                                                                                           u"<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" +
                                                                                           u"<mi>i</mi>" + u"<mo>[+]</mo>" + u"<mi>b</mi>" + u"<mo>-</mo>" +
                                                                                           u"<mi>c</mi>" + u"<mo>≥</mo>" + u".*" + u"</math>" + u"</p>")
                            ->get_Success());
            break;

        case HtmlOfficeMathOutputMode::Text:
            ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(
                            outDocContents, String(u"<p style=\\\"margin-top:0pt; margin-bottom:10pt\\\">") +
                                                u"<span style=\\\"font-family:'Cambria Math'\\\">i[+]b-c≥iM[+]bM-cM </span>" + u"</p>")
                            ->get_Success());
            break;
        }
        //ExEnd
    }

    void ScaleImageToShapeSize(bool scaleImageToShapeSize)
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ScaleImageToShapeSize
        //ExSummary:Shows how to disable the scaling of images to their parent shape dimensions when saving to .html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a shape which contains an image, and then make that shape considerably smaller than the image.
        SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromFile(ImageDir + u"Transparent background logo.png");

        ASSERT_EQ(400, image->get_Size().get_Width());
        ASSERT_EQ(400, image->get_Size().get_Height());

        SharedPtr<Shape> imageShape = builder->InsertImage(image);
        imageShape->set_Width(50);
        imageShape->set_Height(50);

        // Saving a document that contains shapes with images to HTML will create an image file in the local file system
        // for each such shape. The output HTML document will use <image> tags to link to and display these images.
        // When we save the document to HTML, we can pass a SaveOptions object to determine
        // whether to scale all images that are inside shapes to the sizes of their shapes.
        // Setting the "ScaleImageToShapeSize" flag to "true" will shrink every image
        // to the size of the shape that contains it, so that no saved images will be larger than the document requires them to be.
        // Setting the "ScaleImageToShapeSize" flag to "false" will preserve these images' original sizes,
        // which will take up more space in exchange for preserving image quality.
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ScaleImageToShapeSize(scaleImageToShapeSize);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ScaleImageToShapeSize.html", options);

        auto fileInfo = MakeObject<System::IO::FileInfo>(ArtifactsDir + u"HtmlSaveOptions.ScaleImageToShapeSize.001.png");

        //ExEnd
    }

    void ImageFolder()
    {
        //ExStart
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.ExportTextInputFormFieldAsText
        //ExFor:HtmlSaveOptions.ImagesFolder
        //ExSummary:Shows how to specify the folder for storing linked images after saving to .html.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        String imagesDir = System::IO::Path::Combine(ArtifactsDir, u"SaveHtmlWithOptions");

        if (System::IO::Directory::Exists(imagesDir))
        {
            System::IO::Directory::Delete(imagesDir, true);
        }

        System::IO::Directory::CreateDirectory_(imagesDir);

        // Set an option to export form fields as plain text instead of HTML input elements.
        auto options = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        options->set_ExportTextInputFormFieldAsText(true);
        options->set_ImagesFolder(imagesDir);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.SaveHtmlWithOptions.html", options);
        //ExEnd

        ASSERT_TRUE(System::IO::File::Exists(ArtifactsDir + u"HtmlSaveOptions.SaveHtmlWithOptions.html"));
        ASSERT_EQ(9, System::IO::Directory::GetFiles(imagesDir)->get_Length());

        System::IO::Directory::Delete(imagesDir, true);
    }

    //ExStart
    //ExFor:ImageSavingArgs.CurrentShape
    //ExFor:ImageSavingArgs.Document
    //ExFor:ImageSavingArgs.ImageStream
    //ExFor:ImageSavingArgs.IsImageAvailable
    //ExFor:ImageSavingArgs.KeepImageStreamOpen
    //ExSummary:Shows how to involve an image saving callback in an HTML conversion process.
    void ImageSavingCallback()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // When we save the document to HTML, we can pass a SaveOptions object to designate a callback
        // to customize the image saving process.
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ImageSavingCallback(MakeObject<ExHtmlSaveOptions::ImageShapePrinter>());

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ImageSavingCallback.html", options);
    }

    /// <summary>
    /// Prints the properties of each image as the saving process saves it to an image file in the local file system
    /// during the exporting of a document to HTML.
    /// </summary>
    class ImageShapePrinter : public IImageSavingCallback
    {
    public:
        void ImageSaving(SharedPtr<ImageSavingArgs> args) override
        {
            args->set_KeepImageStreamOpen(false);
            ASSERT_TRUE(args->get_IsImageAvailable());

            std::cout << args->get_Document()->get_OriginalFileName().Split(MakeArray<char16_t>({u'\\'}))->LINQ_Last() << " Image #" << ++mImageCount
                      << std::endl;

            auto layoutCollector = MakeObject<LayoutCollector>(args->get_Document());

            std::cout << "\tOn page:\t" << layoutCollector->GetStartPageIndex(args->get_CurrentShape()) << std::endl;
            std::cout << "\tDimensions:\t" << args->get_CurrentShape()->get_Bounds() << std::endl;
            std::cout << String::Format(u"\tAlignment:\t{0}", args->get_CurrentShape()->get_VerticalAlignment()) << std::endl;
            std::cout << String::Format(u"\tWrap type:\t{0}", args->get_CurrentShape()->get_WrapType()) << std::endl;
            std::cout << "Output filename:\t" << args->get_ImageFileName() << "\n" << std::endl;
        }

        ImageShapePrinter() : mImageCount(0)
        {
        }

    private:
        int mImageCount;
    };
    //ExEnd

    void PrettyFormat(bool usePrettyFormat)
    {
        //ExStart
        //ExFor:SaveOptions.PrettyFormat
        //ExSummary:Shows how to enhance the readability of the raw code of a saved .html document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        auto htmlOptions = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        htmlOptions->set_PrettyFormat(usePrettyFormat);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.PrettyFormat.html", htmlOptions);

        // Enabling pretty format makes the raw html code more readable by adding tab stop and new line characters.
        String html = System::IO::File::ReadAllText(ArtifactsDir + u"HtmlSaveOptions.PrettyFormat.html");

        if (usePrettyFormat)
        {
            ASSERT_EQ(String(u"<html>\r\n") + u"\t<head>\r\n" + u"\t\t<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />\r\n" +
                          u"\t\t<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />\r\n" +
                          String::Format(u"\t\t<meta name=\"generator\" content=\"{0} {1}\" />\r\n", BuildVersionInfo::get_Product(),
                                         BuildVersionInfo::get_Version()) +
                          u"\t\t<title></title>\r\n" + u"\t</head>\r\n" + u"\t<body style=\"font-family:'Times New Roman'; font-size:12pt\">\r\n" +
                          u"\t\t<div>\r\n" + u"\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">\r\n" + u"\t\t\t\t<span>Hello world!</span>\r\n" +
                          u"\t\t\t</p>\r\n" + u"\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">\r\n" +
                          u"\t\t\t\t<span style=\"-aw-import:ignore\">&#xa0;</span>\r\n" + u"\t\t\t</p>\r\n" + u"\t\t</div>\r\n" + u"\t</body>\r\n</html>\r\n",
                      html);
        }
        else
        {
            ASSERT_EQ(String(u"<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />") +
                          u"<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />" +
                          String::Format(u"<meta name=\"generator\" content=\"{0} {1}\" /><title></title></head>", BuildVersionInfo::get_Product(),
                                         BuildVersionInfo::get_Version()) +
                          u"<body style=\"font-family:'Times New Roman'; font-size:12pt\">" +
                          u"<div><p style=\"margin-top:0pt; margin-bottom:0pt\"><span>Hello world!</span></p>" +
                          u"<p style=\"margin-top:0pt; margin-bottom:0pt\"><span style=\"-aw-import:ignore\">&#xa0;</span></p></div></body></html>",
                      html);
        }
        //ExEnd
    }
};

} // namespace ApiExamples
