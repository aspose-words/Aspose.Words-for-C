#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Layout/Public/LayoutCollector.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/BuildVersionInfo.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldToc.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Saving/CssStyleSheetType.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentSplitCriteria.h>
#include <Aspose.Words.Cpp/Model/Saving/ExportListLabels.h>
#include <Aspose.Words.Cpp/Model/Saving/FontSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlElementSizeOutputMode.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlMetafileFormat.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlOfficeMathOutputMode.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlVersion.h>
#include <Aspose.Words.Cpp/Model/Saving/IFontSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/IImageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/PaperSize.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <drawing/rectangle_f.h>
#include <system/array.h>
#include <system/enum.h>
#include <system/enum_helpers.h>
#include <system/exceptions.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/io/path.h>
#include <system/io/search_option.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/primitive_types.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/test_tools.h>
#include <system/text/encoding.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

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
    void ExportPageMargins(SaveFormat saveFormat)
    {
        auto doc = MakeObject<Document>(MyDir + u"TextBoxes.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_SaveFormat(saveFormat);
        saveOptions->set_ExportPageMargins(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportPageMargins" + FileFormatUtil::SaveFormatToExtension(saveFormat), saveOptions);
    }

    void ExportOfficeMath(SaveFormat saveFormat, HtmlOfficeMathOutputMode outputMode)
    {
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_OfficeMathOutputMode(outputMode);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportOfficeMath" + FileFormatUtil::SaveFormatToExtension(saveFormat), saveOptions);
    }

    void ExportTextBoxAsSvg(SaveFormat saveFormat, bool isTextBoxAsSvg)
    {
        ArrayPtr<String> dirFiles;

        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> textbox = builder->InsertShape(ShapeType::TextBox, 300, 100);
        builder->MoveTo(textbox->get_FirstParagraph());
        builder->Write(u"Hello world!");

        auto saveOptions = MakeObject<HtmlSaveOptions>(saveFormat);
        saveOptions->set_ExportTextBoxAsSvg(isTextBoxAsSvg);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportTextBoxAsSvg" + FileFormatUtil::SaveFormatToExtension(saveFormat), saveOptions);

        switch (saveFormat)
        {
        case SaveFormat::Html:
            dirFiles = System::IO::Directory::GetFiles(ArtifactsDir, u"HtmlSaveOptions.ExportTextBoxAsSvg.001.png", System::IO::SearchOption::AllDirectories);
            ASSERT_EQ(0, dirFiles->get_Length());
            return;

        case SaveFormat::Epub:
            dirFiles = System::IO::Directory::GetFiles(ArtifactsDir, u"HtmlSaveOptions.ExportTextBoxAsSvg.001.png", System::IO::SearchOption::AllDirectories);
            ASSERT_EQ(0, dirFiles->get_Length());
            return;

        case SaveFormat::Mhtml:
            dirFiles = System::IO::Directory::GetFiles(ArtifactsDir, u"HtmlSaveOptions.ExportTextBoxAsSvg.001.png", System::IO::SearchOption::AllDirectories);
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

        doc->Save(ArtifactsDir + String::Format(u"HtmlSaveOptions.ControlListLabelsExport.html"), saveOptions);
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
        // Assert that default value is true for HTML and false for MHTML and EPUB
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

        ArrayPtr<String> imageFiles = System::IO::Directory::GetFiles(
            ArtifactsDir + u"Resources/", u"HtmlSaveOptions.ExternalResourceSavingConfig*.png", System::IO::SearchOption::AllDirectories);
        ASSERT_EQ(8, imageFiles->get_Length());

        ArrayPtr<String> fontFiles = System::IO::Directory::GetFiles(
            ArtifactsDir + u"Resources/", u"HtmlSaveOptions.ExternalResourceSavingConfig*.ttf", System::IO::SearchOption::AllDirectories);
        ASSERT_EQ(10, fontFiles->get_Length());

        ArrayPtr<String> cssFiles = System::IO::Directory::GetFiles(
            ArtifactsDir + u"Resources/", u"HtmlSaveOptions.ExternalResourceSavingConfig*.css", System::IO::SearchOption::AllDirectories);
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
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        String fontsFolder = ArtifactsDir + u"HtmlSaveOptions.ExportFonts.Resources";
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

        ArrayPtr<String> a = System::IO::Directory::GetFiles(ArtifactsDir + u"Resources", u"HtmlSaveOptions.ResourceFolderPriority.001.jpeg",
                                                                             System::IO::SearchOption::AllDirectories);
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

        builder->get_Document()->Save(ArtifactsDir + u"HtmlSaveOptions.SvgMetafileFormat.html",
                                      [&]
                                      {
                                          auto tmp_0 = MakeObject<HtmlSaveOptions>();
                                          tmp_0->set_MetafileFormat(HtmlMetafileFormat::Png);
                                          return tmp_0;
                                      }());
    }

    void PngMetafileFormat()
    {
        auto builder = MakeObject<DocumentBuilder>();

        builder->Write(u"Here is an Png image: ");
        builder->InsertHtml(u"<svg height='210' width='500'>\r\n                    <polygon points='100,10 40,198 190,78 10,78 160,198' \r\n                  "
                            u"      style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n                  </svg> ");

        builder->get_Document()->Save(ArtifactsDir + u"HtmlSaveOptions.PngMetafileFormat.html",
                                      [&]
                                      {
                                          auto tmp_1 = MakeObject<HtmlSaveOptions>();
                                          tmp_1->set_MetafileFormat(HtmlMetafileFormat::Png);
                                          return tmp_1;
                                      }());
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

        builder->get_Document()->Save(ArtifactsDir + u"HtmlSaveOptions.EmfOrWmfMetafileFormat.html",
                                      [&]
                                      {
                                          auto tmp_2 = MakeObject<HtmlSaveOptions>();
                                          tmp_2->set_MetafileFormat(HtmlMetafileFormat::EmfOrWmf);
                                          return tmp_2;
                                      }());
    }

    void CssClassNamesPrefix()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.CssClassNamePrefix
        //ExSummary:Shows how to specifies a prefix which is added to all CSS class names.
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_CssStyleSheetType(CssStyleSheetType::Embedded);
        saveOptions->set_CssClassNamePrefix(u"myprefix-");

        // The prefix will be found before CSS element names in the embedded stylesheet
        doc->Save(ArtifactsDir + u"HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
        //ExEnd
    }

    void CssClassNamesNotValidPrefix()
    {
        auto saveOptions = MakeObject<HtmlSaveOptions>();

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<void()> _local_func_0 = [&saveOptions]()
        {
            saveOptions->set_CssClassNamePrefix(u"@%-");
        };

        ASSERT_THROW(_local_func_0(), System::ArgumentException) << "The class name prefix must be a valid CSS identifier.";
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
        //ExSummary:Shows how to split a document into several html documents by heading levels.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert headings of levels 1 - 3
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

        // Create a HtmlSaveOptions object and set the split criteria to "HeadingParagraph", meaning that the document
        // will be split into parts at the beginning of every paragraph of a "Heading" style, and each part will be saved as a separate document
        // Also, we will set the DocumentSplitHeadingLevel to 2, which will split the document only at headings that have levels from 1 to 2
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_DocumentSplitCriteria(DocumentSplitCriteria::HeadingParagraph);
        options->set_DocumentSplitHeadingLevel(2);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.HeadingLevels.html", options);
        //ExEnd
    }

    void NegativeIndent()
    {
        //ExStart
        //ExFor:HtmlElementSizeOutputMode
        //ExFor:HtmlSaveOptions.AllowNegativeIndent
        //ExFor:HtmlSaveOptions.TableWidthOutputMode
        //ExSummary:Shows how to preserve negative indents in the output .html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a table and give it a negative value for its indent, effectively pushing it out of the left page boundary
        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Cell 1");
        builder->InsertCell();
        builder->Write(u"Cell 2");
        builder->EndTable();
        table->set_LeftIndent(-36);
        table->set_PreferredWidth(PreferredWidth::FromPoints(144));

        // When saving to .html, this indent will only be preserved if we set this flag
        auto options = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        options->set_AllowNegativeIndent(true);
        options->set_TableWidthOutputMode(HtmlElementSizeOutputMode::RelativeOnly);

        // The first cell with "Cell 1" will not be visible in the output
        doc->Save(ArtifactsDir + u"HtmlSaveOptions.NegativeIndent.html", options);
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
        //ExSummary:Shows how to set folders and folder aliases for externally saved resources when saving to html.
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

        // Configure a SaveOptions object to export fonts to separate files, in a manner specified by a custom callback.
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportFontResources(true);
        options->set_FontSavingCallback(MakeObject<ExHtmlSaveOptions::HandleFontSaving>());

        // The callback will export .ttf files and saved alongside the output document.
        doc->Save(ArtifactsDir + u"HtmlSaveOptions.SaveExportedFonts.html", options);

        std::function<bool(String s)> endsWithTtf = [](String s)
        {
            return s.EndsWith(u".ttf");
        };

        for (String t : System::Array<String>::FindAll(System::IO::Directory::GetFiles(ArtifactsDir), endsWithTtf))
        {
            std::cout << t << std::endl;
        }

        ASSERT_EQ(10, System::Array<String>::FindAll(System::IO::Directory::GetFiles(ArtifactsDir), endsWithTtf)->get_Length());
        //ExSkip
    }

    /// <summary>
    /// Prints information about exported fonts, and saves them alongside their output .html.
    /// </summary>
    class HandleFontSaving : public IFontSavingCallback
    {
    public:
        void FontSaving(SharedPtr<FontSavingArgs> args) override
        {
            // Print information for every font resource that is about to be saved.
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

            // The source document can also be accessed from here.
            ASSERT_TRUE(args->get_Document()->get_OriginalFileName().EndsWith(u"Rendering.docx"));

            ASSERT_TRUE(args->get_IsExportNeeded());
            ASSERT_TRUE(args->get_IsSubsettingNeeded());

            // There are two ways of saving an exported font.
            // 1 -  Save it to a local file system location determined by a filename:
            args->set_FontFileName(args->get_OriginalFileName().Split(MakeArray<char16_t>({System::IO::Path::DirectorySeparatorChar}))->LINQ_Last());

            // 2 -  Save it to a stream.
            args->set_FontStream(MakeObject<System::IO::FileStream>(
                ArtifactsDir + args->get_OriginalFileName().Split(MakeArray<char16_t>({System::IO::Path::DirectorySeparatorChar}))->LINQ_Last(),
                System::IO::FileMode::Create));
            ASSERT_FALSE(args->get_KeepFontStreamOpen());
        }
    };
    //ExEnd

    void HtmlVersion_()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.#ctor(SaveFormat)
        //ExFor:HtmlSaveOptions.ExportXhtmlTransitional
        //ExFor:HtmlSaveOptions.HtmlVersion
        //ExFor:HtmlVersion
        //ExSummary:Shows how to set a saved .html document to a specific version.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Save the document to a .html file of the XHTML 1.0 Transitional standard
        auto options = MakeObject<HtmlSaveOptions>(SaveFormat::Html);
        options->set_HtmlVersion(HtmlVersion::Xhtml);
        options->set_ExportXhtmlTransitional(true);
        options->set_PrettyFormat(true);

        // The DOCTYPE declaration at the top of this document will indicate the html version we chose
        doc->Save(ArtifactsDir + u"HtmlSaveOptions.HtmlVersion.html", options);
        //ExEnd
    }

    void EpubHeadings()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.EpubNavigationMapLevel
        //ExSummary:Shows the relationship between heading levels and the Epub navigation panel.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert headings of levels 1 - 3
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

        // Epub readers normally treat paragraphs with "Heading" styles as anchors for a table of contents-style navigation pane
        // We set a maximum heading level above which headings will not be registered by the reader as navigation points with
        // a HtmlSaveOptions object and its EpubNavigationLevel attribute
        // Our document has headings of levels 1 to 3,
        // but our output epub will only place level 1 and 2 headings in the table of contents
        auto options = MakeObject<HtmlSaveOptions>(SaveFormat::Epub);
        options->set_EpubNavigationMapLevel(2);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.EpubHeadings.epub", options);
        //ExEnd
    }

    void Doc2EpubSaveOptions()
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
        //ExSummary:Shows how to specify saving options while converting a document to .epub.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Specify encoding for a document that we will save with a SaveOptions object.
        auto saveOptions = MakeObject<HtmlSaveOptions>();
        saveOptions->set_SaveFormat(SaveFormat::Epub);
        saveOptions->set_Encoding(System::Text::Encoding::get_UTF8());

        // By default, an output .epub document will have all the contents in one HTML part.
        // A split criteria allows us to segment the document into several HTML parts.
        // We will set the criteria to split the document at heading paragraphs.
        // This is useful for readers which cannot read HTML files greater than a certain size.
        saveOptions->set_DocumentSplitCriteria(DocumentSplitCriteria::HeadingParagraph);

        // Specify that we want to export document properties.
        saveOptions->set_ExportDocumentProperties(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.Doc2EpubSaveOptions.epub", saveOptions);
        //ExEnd
    }

    void ContentIdUrls()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportCidUrlsForMhtmlResources
        //ExSummary:Shows how to enable content IDs for output MHTML documents.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Setting this flag will replace "Content-Location" tags with "Content-ID" tags for each resource from the input document
        // The file names that were next to each "Content-Location" tag are re-purposed as content IDs
        auto options = MakeObject<HtmlSaveOptions>(SaveFormat::Mhtml);
        options->set_ExportCidUrlsForMhtmlResources(true);
        options->set_CssStyleSheetType(CssStyleSheetType::External);
        options->set_ExportFontResources(true);
        options->set_PrettyFormat(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ContentIdUrls.mht", options);
        //ExEnd
    }

    void DropDownFormField()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportDropDownFormFieldAsText
        //ExSummary:Shows how to get drop down combo box form fields to blend in with paragraph text when saving to html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to insert a combo box with the value "Two" selected
        builder->InsertComboBox(u"MyComboBox", MakeArray<String>({u"One", u"Two", u"Three"}), 1);

        // When converting to .html, drop down combo boxes will be converted to select/option tags to preserve their functionality
        // If we want to freeze a combo box at its current selected value and convert it into plain text, we can do so with this flag
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportDropDownFormFieldAsText(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.DropDownFormField.html", options);
        //ExEnd
    }

    void ExportBase64()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportFontsAsBase64
        //ExFor:HtmlSaveOptions.ExportImagesAsBase64
        //ExSummary:Shows how to save a .html document with resources embedded inside it.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // By default, when converting a document with images to .html, resources such as images will be linked to in external files
        // We can set these flags to embed resources inside the output .html instead, reducing the number of files created during the conversion
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportFontsAsBase64(true);
        options->set_ExportImagesAsBase64(true);
        options->set_PrettyFormat(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportBase64.html", options);
        //ExEnd
    }

    void ExportLanguageInformation()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportLanguageInformation
        //ExSummary:Shows how to preserve language information when saving to .html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use the builder to write text in more than one language
        builder->get_Font()->set_LocaleId(2057);
        // en-GB
        builder->Writeln(u"Hello world!");

        builder->get_Font()->set_LocaleId(1049);
        // ru-RU
        builder->Write(u"Привет, мир!");

        // Normally, when saving a document with more than one proofing language to .html,
        // only the text content is preserved with no traces of any other languages
        // Saving with a HtmlSaveOptions object with this flag set will add "lang" attributes to spans
        // in places where other proofing languages were used
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportLanguageInformation(true);
        options->set_PrettyFormat(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportLanguageInformation.html", options);
        //ExEnd
    }

    void List_()
    {
        //ExStart
        //ExFor:ExportListLabels
        //ExFor:HtmlSaveOptions.ExportListLabels
        //ExSummary:Shows how to export an indented list to .html as plain text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use the builder to insert a list
        SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(ListTemplate::NumberDefault);
        builder->get_ListFormat()->set_List(list);

        builder->Writeln(u"List item 1.");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"List item 2.");
        builder->get_ListFormat()->ListIndent();
        builder->Write(u"List item 3.");

        // When we save this to .html, normally our list will be represented by <li> tags
        // We can set this flag to have lists as plain text instead
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportListLabels(ExportListLabels::AsInlineText);
        options->set_PrettyFormat(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.List.html", options);
        //ExEnd
    }

    void ExportPageMargins2()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportPageMargins
        //ExSummary:Shows how to show out-of-bounds objects in output .html documents.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a builder to insert a shape with no wrapping
        SharedPtr<Shape> shape = builder->InsertShape(ShapeType::Cube, 200, 200);

        shape->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);
        shape->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);
        shape->set_WrapType(WrapType::None);

        // Negative values for shape position may cause the shape to go out of page bounds
        // If we export this to .html, the shape will be truncated
        shape->set_Left(-150);

        // We can avoid that and have the entire shape be visible by setting this flag
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportPageMargins(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportPageMargins2.html", options);
        //ExEnd
    }

    void ExportPageSetup()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportPageSetup
        //ExSummary:Shows how to preserve section structure/page setup information when saving to html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a DocumentBuilder to insert two sections with text
        builder->Writeln(u"Section 1");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Writeln(u"Section 2");

        // Change dimensions and paper size of first section
        SharedPtr<PageSetup> pageSetup = doc->get_Sections()->idx_get(0)->get_PageSetup();
        pageSetup->set_TopMargin(36.0);
        pageSetup->set_BottomMargin(36.0);
        pageSetup->set_PaperSize(PaperSize::A5);

        // Section structure and pagination are normally lost when when converting to .html
        // We can create an HtmlSaveOptions object with the ExportPageSetup flag set to true
        // to preserve the section structure in <div> tags and page dimensions in the output document's CSS
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportPageSetup(true);
        options->set_PrettyFormat(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportPageSetup.html", options);
        //ExEnd
    }

    void RelativeFontSize()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportRelativeFontSize
        //ExSummary:Shows how to use relative font sizes when saving to .html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a builder to write some text in various sizes
        builder->Writeln(u"Default font size, ");
        builder->get_Font()->set_Size(24.0);
        builder->Writeln(u"2x default font size,");
        builder->get_Font()->set_Size(96);
        builder->Write(u"8x default font size");

        // We can save font sizes as ratios of the default size, which will be 12 in this case
        // If we use an input .html, this size can be set with the AbsSize {font-size:12pt} tag
        // The ExportRelativeFontSize will enable this feature
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportRelativeFontSize(true);
        options->set_PrettyFormat(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.RelativeFontSize.html", options);
        //ExEnd
    }

    void ExportTextBox()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportTextBoxAsSvg
        //ExSummary:Shows how to export text boxes as scalable vector graphics.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a DocumentBuilder to insert a text box and give it some text content
        SharedPtr<Shape> textBox = builder->InsertShape(ShapeType::TextBox, 100.0, 60.0);
        builder->MoveTo(textBox->get_FirstParagraph());
        builder->Write(u"My text box");

        // Normally, all shapes such as the text box we placed are exported to .html as external images linked by the .html document
        // We can save with an HtmlSaveOptions object with the ExportTextBoxAsSvg set to true to save text boxes as <svg> tags,
        // which will cause no linked images to be saved and will make the inner text selectable
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportTextBoxAsSvg(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportTextBox.html", options);
        //ExEnd
    }

    void RoundTripInformation()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportRoundtripInformation
        //ExSummary:Shows how to preserve hidden elements when converting to .html.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // When converting a document to .html, some elements such as hidden bookmarks, original shape positions,
        // or footnotes will be either removed or converted to plain text and effectively be lost
        // Saving with a HtmlSaveOptions object with ExportRoundtripInformation set to true will preserve these elements
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportRoundtripInformation(true);
        options->set_PrettyFormat(true);

        // These elements will have tags that will start with "-aw", such as "-aw-import" or "-aw-left-pos"
        doc->Save(ArtifactsDir + u"HtmlSaveOptions.RoundTripInformation.html", options);
        //ExEnd
    }

    void ExportTocPageNumbers()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportTocPageNumbers
        //ExSummary:Shows how to display page numbers when saving a document with a table of contents to .html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a table of contents
        auto fieldToc = System::DynamicCast<FieldToc>(builder->InsertField(FieldType::FieldTOC, true));

        // Populate the document with paragraphs of a "Heading" style that the table of contents will pick up
        builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(u"Heading 1"));
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Entry 1");
        builder->Writeln(u"Entry 2");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Entry 3");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Entry 4");

        // Our headings span several pages, and those page numbers will be displayed by the TOC at the top of the document
        fieldToc->UpdatePageNumbers();
        doc->UpdateFields();

        // These page numbers are normally omitted since .html has no pagination, but we can still have them displayed
        // if we save with a HtmlSaveOptions object with the ExportTocPageNumbers set to true
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportTocPageNumbers(true);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ExportTocPageNumbers.html", options);
        //ExEnd
    }

    void FontSubsetting()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.FontResourcesSubsettingSizeThreshold
        //ExSummary:Shows how to work with font subsetting.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a DocumentBuilder to insert text with several fonts
        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(u"Hello world!");
        builder->get_Font()->set_Name(u"Times New Roman");
        builder->Writeln(u"Hello world!");
        builder->get_Font()->set_Name(u"Courier New");
        builder->Writeln(u"Hello world!");

        // When saving to .html, font subsetting fully applies by default, meaning that when we export fonts with our file,
        // the symbols not used by our document are not represented by the exported fonts, which cuts down file size dramatically
        // Font files of a file size larger than FontResourcesSubsettingSizeThreshold get subsetted, so a value of 0 will apply default full subsetting
        // Setting the value to something large will fully suppress subsetting, which could result in large font files that cover every glyph
        String fontsFolder = ArtifactsDir + u"HtmlSaveOptions.FontSubsetting.Fonts";
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ExportFontResources(true);
        options->set_FontsFolder(fontsFolder);
        options->set_FontResourcesSubsettingSizeThreshold(std::numeric_limits<int>::max());

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.FontSubsetting.html", options);
        //ExEnd
    }

    void MetafileFormat()
    {
        //ExStart
        //ExFor:HtmlMetafileFormat
        //ExFor:HtmlSaveOptions.MetafileFormat
        //ExSummary:Shows how to set a meta file in a different format.
        // Create a document from an html string
        String html =
            u"<html>\r\n                    <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>\r\n                        "
            u"<text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>\r\n                    </svg>\r\n                </html>";

        auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(html)));

        // This document contains a <svg> element in the form of text,
        // which by default will be saved as a linked external .png when we save the document as html
        // We can save with a HtmlSaveOptions object with this flag set to preserve the <svg> tag
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_MetafileFormat(HtmlMetafileFormat::Svg);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.MetafileFormat.html", options);
        //ExEnd
    }

    void OfficeMathOutputMode()
    {
        //ExStart
        //ExFor:HtmlOfficeMathOutputMode
        //ExFor:HtmlSaveOptions.OfficeMathOutputMode
        //ExSummary:Shows how to control the way how OfficeMath objects are exported to .html.
        // Open a document that contains OfficeMath objects
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");

        // Create a HtmlSaveOptions object and configure it to export OfficeMath objects as images
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_OfficeMathOutputMode(HtmlOfficeMathOutputMode::Image);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.OfficeMathOutputMode.html", options);
        //ExEnd
    }

    void ScaleImageToShapeSize()
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ScaleImageToShapeSize
        //ExSummary:Shows how to disable the scaling of images to their parent shape dimensions when saving to .html.
        // Open a document which contains shapes with images
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // By default, images inside shapes get scaled to the size of their shapes while the document gets
        // converted to .html, reducing image file size
        // We can save the document with a HtmlSaveOptions with ScaleImageToShapeSize set to false to prevent the scaling
        // and preserve the full quality and file size of the linked images
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ScaleImageToShapeSize(false);

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ScaleImageToShapeSize.html", options);
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

        // Set a directory where images will be saved to, then ensure that it exists, and is empty.
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
    //ExSummary:Shows how to involve an image saving callback in an .html conversion process.
    void ImageSavingCallback()
    {
        // Open a document which contains shapes with images
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Create a HtmlSaveOptions object with a custom image saving callback that will print image information
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_ImageSavingCallback(MakeObject<ExHtmlSaveOptions::ImageShapePrinter>());

        doc->Save(ArtifactsDir + u"HtmlSaveOptions.ImageSavingCallback.html", options);
    }

private:
    /// <summary>
    /// Prints information of all images that are about to be saved from within a document to image files
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
            std::cout << "\tDimensions:\t" << System::ObjectExt::ToString(args->get_CurrentShape()->get_Bounds()) << std::endl;
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

public:
    void PrettyFormat(bool usePrettyFormat)
    {
        //ExStart
        //ExFor:SaveOptions.PrettyFormat
        //ExSummary:Shows how to enhance the readability of the raw code of a saved .html document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // Enable pretty format via a SaveOptions object, then save the document in .html to the local file system.
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
                          u"\t\t<title>\r\n" + u"\t\t</title>\r\n" + u"\t</head>\r\n" +
                          u"\t<body style=\"font-family:'Times New Roman'; font-size:12pt\">\r\n" + u"\t\t<div>\r\n" +
                          u"\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">\r\n" + u"\t\t\t\t<span>Hello world!</span>\r\n" + u"\t\t\t</p>\r\n" +
                          u"\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">\r\n" + u"\t\t\t\t<span style=\"-aw-import:ignore\">&#xa0;</span>\r\n" +
                          u"\t\t\t</p>\r\n" + u"\t\t</div>\r\n" + u"\t</body>\r\n</html>",
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
