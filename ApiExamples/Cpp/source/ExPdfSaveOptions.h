#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/CertificateHolder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Document/WarningSource.h>
#include <Aspose.Words.Cpp/Model/Document/WarningType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Properties/CustomDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Properties/DocumentProperty.h>
#include <Aspose.Words.Cpp/Model/Saving/ColorMode.h>
#include <Aspose.Words.Cpp/Model/Saving/Dml3DEffectsRenderingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/DmlEffectsRenderingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/DmlRenderingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/DownsampleOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/EmfPlusDualRenderingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/HeaderFooterBookmarksExportMode.h>
#include <Aspose.Words.Cpp/Model/Saving/MetafileRenderingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/MetafileRenderingOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OutlineOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfCompliance.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfCustomPropertiesExport.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfDigitalSignatureDetails.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfDigitalSignatureHashAlgorithm.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfDigitalSignatureTimestampSettings.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfImageColorSpaceExportMode.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfImageCompression.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfPageMode.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfTextCompression.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfZoomBehavior.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Settings/MultiplePagesType.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <drawing/image.h>
#include <system/collections/list.h>
#include <system/date_time.h>
#include <system/enum.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/globalization/culture_info.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/test_tools.h>
#include <system/timespan.h>
#include <testing/test_predicates.h>
#include <gtest/gtest-spi.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

namespace ApiExamples {
class ExPdfSaveOptions;
}

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace ApiExamples {

class ExPdfSaveOptions : public ApiExampleBase
{
public:
    void CreateMissingOutlineLevels()
    {
        //ExStart
        //ExFor:OutlineOptions.CreateMissingOutlineLevels
        //ExFor:ParagraphFormat.IsHeading
        //ExFor:PdfSaveOptions.OutlineOptions
        //ExFor:PdfSaveOptions.SaveFormat
        //ExSummary:Shows how to create PDF document outline entries for headings.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create TOC entries
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);
        ASSERT_TRUE(builder->get_ParagraphFormat()->get_IsHeading());

        builder->Writeln(u"Heading 1");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading4);

        builder->Writeln(u"Heading 1.1.1.1");
        builder->Writeln(u"Heading 1.1.1.2");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading9);

        builder->Writeln(u"Heading 1.1.1.1.1.1.1.1.1");
        builder->Writeln(u"Heading 1.1.1.1.1.1.1.1.2");

        // Create "PdfSaveOptions" with some mandatory parameters
        // "HeadingsOutlineLevels" specifies how many levels of headings to include in the document outline
        // "CreateMissingOutlineLevels" determining whether to create missing heading levels
        auto pdfSaveOptions = MakeObject<PdfSaveOptions>();
        pdfSaveOptions->get_OutlineOptions()->set_HeadingsOutlineLevels(9);
        pdfSaveOptions->get_OutlineOptions()->set_CreateMissingOutlineLevels(true);
        pdfSaveOptions->set_SaveFormat(SaveFormat::Pdf);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.CreateMissingOutlineLevels.pdf", pdfSaveOptions);
        //ExEnd

    }

    void TableHeadingOutlines()
    {
        //ExStart
        //ExFor:OutlineOptions.CreateOutlinesForHeadingsInTables
        //ExSummary:Shows how to create PDF document outline entries for headings inside tables.
        // Create a blank document and insert a table with a heading-style text inside it
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartTable();
        builder->InsertCell();
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);
        builder->Write(u"Heading 1");
        builder->EndRow();
        builder->InsertCell();
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Normal);
        builder->Write(u"Cell 1");
        builder->EndTable();

        // Create a PdfSaveOptions object that, when saving to .pdf with it, creates entries in the document outline for heading levels 1-9,
        // and make sure headings inside tables are registered by the outline also
        auto pdfSaveOptions = MakeObject<PdfSaveOptions>();
        pdfSaveOptions->get_OutlineOptions()->set_HeadingsOutlineLevels(9);
        pdfSaveOptions->get_OutlineOptions()->set_CreateOutlinesForHeadingsInTables(true);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.TableHeadingOutlines.pdf", pdfSaveOptions);
        //ExEnd

    }

    void UpdateFields(bool doUpdateFields)
    {
        //ExStart
        //ExFor:PdfSaveOptions.Clone
        //ExFor:SaveOptions.UpdateFields
        //ExSummary:Shows how to update fields before saving into a PDF document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert two pages of text, including two fields that will need to be updated to display an accurate value
        builder->Write(u"Page ");
        builder->InsertField(u"PAGE", u"");
        builder->Write(u" of ");
        builder->InsertField(u"NUMPAGES", u"");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Hello World!");

        auto options = MakeObject<PdfSaveOptions>();
        options->set_UpdateFields(doUpdateFields);

        // PdfSaveOptions objects can be cloned
        ASPOSE_ASSERT_NS(options, options->Clone());

        doc->Save(ArtifactsDir + u"PdfSaveOptions.UpdateFields.pdf", options);
        //ExEnd

    }

    void ImageCompression()
    {
        //ExStart
        //ExFor:PdfSaveOptions.Compliance
        //ExFor:PdfSaveOptions.ImageCompression
        //ExFor:PdfSaveOptions.ImageColorSpaceExportMode
        //ExFor:PdfSaveOptions.JpegQuality
        //ExFor:PdfImageCompression
        //ExFor:PdfCompliance
        //ExFor:PdfImageColorSpaceExportMode
        //ExSummary:Shows how to save images to PDF using JPEG encoding to decrease file size.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        auto pdfSaveOptions = MakeObject<PdfSaveOptions>();
        pdfSaveOptions->set_ImageCompression(PdfImageCompression::Jpeg);
        pdfSaveOptions->get_DownsampleOptions()->set_DownsampleImages(false);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.ImageCompression.pdf", pdfSaveOptions);

        auto pdfSaveOptionsA1B = MakeObject<PdfSaveOptions>();
        pdfSaveOptionsA1B->set_Compliance(PdfCompliance::PdfA1b);
        pdfSaveOptionsA1B->set_ImageCompression(PdfImageCompression::Jpeg);
        pdfSaveOptionsA1B->get_DownsampleOptions()->set_DownsampleImages(false);
        // Use JPEG compression at 50% quality to reduce file size
        pdfSaveOptionsA1B->set_JpegQuality(100);
        pdfSaveOptionsA1B->set_ImageColorSpaceExportMode(PdfImageColorSpaceExportMode::SimpleCmyk);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.ImageCompression.PDF_A_1_B.pdf", pdfSaveOptionsA1B);

        auto pdfSaveOptionsA1A = MakeObject<PdfSaveOptions>();
        pdfSaveOptionsA1A->set_Compliance(PdfCompliance::PdfA1a);
        pdfSaveOptionsA1A->set_ExportDocumentStructure(true);
        pdfSaveOptionsA1A->set_ImageCompression(PdfImageCompression::Jpeg);
        pdfSaveOptionsA1A->get_DownsampleOptions()->set_DownsampleImages(false);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.ImageCompression.PDF_A_1_A.pdf", pdfSaveOptionsA1A);
        //ExEnd

    }

    void DownsampleOptions_()
    {
        //ExStart
        //ExFor:DownsampleOptions
        //ExFor:DownsampleOptions.DownsampleImages
        //ExFor:DownsampleOptions.Resolution
        //ExFor:DownsampleOptions.ResolutionThreshold
        //ExFor:PdfSaveOptions.DownsampleOptions
        //ExSummary:Shows how to change the resolution of images in output pdf documents.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Create a SaveOptions object, verify its default image downsampling settings,
        // and then convert the document to .pdf with it.
        auto options = MakeObject<PdfSaveOptions>();

        ASSERT_TRUE(options->get_DownsampleOptions()->get_DownsampleImages());
        ASSERT_EQ(220, options->get_DownsampleOptions()->get_Resolution());

        doc->Save(ArtifactsDir + u"PdfSaveOptions.DownsampleOptions.Default.pdf", options);

        // Set the output resolution to a lower value, then set a downsampling resolution threshold
        // which will prevent any images with a resolution of less than 128 from being downsampled.
        options->get_DownsampleOptions()->set_Resolution(36);
        options->get_DownsampleOptions()->set_ResolutionThreshold(128);

        // Only the first two images from the document will be downsampled at this stage.
        doc->Save(ArtifactsDir + u"PdfSaveOptions.DownsampleOptions.LowerResolution.pdf", options);
        //ExEnd
    }

    void ColorRendering()
    {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExFor:ColorMode
        //ExFor:FixedPageSaveOptions.ColorMode
        //ExSummary:Shows how change image color with save options property.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        // Configure PdfSaveOptions to save every image in the input document in greyscale during conversion
        auto pdfSaveOptions = MakeObject<PdfSaveOptions>();
        pdfSaveOptions->set_ColorMode(ColorMode::Grayscale);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.ColorRendering.pdf", pdfSaveOptions);
        //ExEnd

    }

    void WindowsBarPdfTitle()
    {
        //ExStart
        //ExFor:PdfSaveOptions.DisplayDocTitle
        //ExSummary:Shows how to display title of the document as title bar.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        doc->get_BuiltInDocumentProperties()->set_Title(u"Windows bar pdf title");

        auto pdfSaveOptions = MakeObject<PdfSaveOptions>();
        pdfSaveOptions->set_DisplayDocTitle(true);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.WindowsBarPdfTitle.pdf", pdfSaveOptions);
        //ExEnd

    }

    void MemoryOptimization()
    {
        //ExStart
        //ExFor:SaveOptions.CreateSaveOptions(SaveFormat)
        //ExFor:SaveOptions.MemoryOptimization
        //ExSummary:Shows an option to optimize memory consumption when you work with large documents.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // When set to true it will improve document memory footprint but will add extra time to processing
        SharedPtr<SaveOptions> saveOptions = SaveOptions::CreateSaveOptions(SaveFormat::Pdf);
        saveOptions->set_MemoryOptimization(true);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.MemoryOptimization.pdf", saveOptions);
        //ExEnd
    }

    void EscapeUri(String uri, String result, bool isEscaped)
    {
        //ExStart
        //ExFor:PdfSaveOptions.EscapeUri
        //ExFor:PdfSaveOptions.OpenHyperlinksInNewWindow
        //ExSummary:Shows how to escape hyperlinks in the document.
        auto builder = MakeObject<DocumentBuilder>();
        builder->InsertHyperlink(u"Testlink", uri, false);

        // Set this property to false if you are sure that hyperlinks in document's model are already escaped
        auto options = MakeObject<PdfSaveOptions>();
        options->set_EscapeUri(isEscaped);
        options->set_OpenHyperlinksInNewWindow(true);

        builder->get_Document()->Save(ArtifactsDir + u"PdfSaveOptions.EscapedUri.pdf", options);
        //ExEnd

    }

    //ExStart
    //ExFor:MetafileRenderingMode
    //ExFor:MetafileRenderingOptions
    //ExFor:MetafileRenderingOptions.EmulateRasterOperations
    //ExFor:MetafileRenderingOptions.RenderingMode
    //ExFor:IWarningCallback
    //ExFor:FixedPageSaveOptions.MetafileRenderingOptions
    //ExSummary:Shows added fallback to bitmap rendering and changing type of warnings about unsupported metafile records.
    void HandleBinaryRasterWarnings()
    {
        auto doc = MakeObject<Document>(MyDir + u"WMF with image.docx");

        auto metafileRenderingOptions = MakeObject<MetafileRenderingOptions>();
        metafileRenderingOptions->set_EmulateRasterOperations(false);
        metafileRenderingOptions->set_RenderingMode(MetafileRenderingMode::VectorWithFallback);

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words
        // renders this metafile to a bitmap
        auto callback = MakeObject<ExPdfSaveOptions::HandleDocumentWarnings>();
        doc->set_WarningCallback(callback);

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_MetafileRenderingOptions(metafileRenderingOptions);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.HandleBinaryRasterWarnings.pdf", saveOptions);

        ASSERT_EQ(1, callback->Warnings->get_Count());
        ASSERT_EQ(u"'R2_XORPEN' binary raster operation is partly supported.", callback->Warnings->idx_get(0)->get_Description());
    }

    class HandleDocumentWarnings : public IWarningCallback
    {
    public:
        SharedPtr<WarningInfoCollection> Warnings;

        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        void Warning(SharedPtr<WarningInfo> info) override
        {
            // For now, type of warnings about unsupported metafile records changed from
            // DataLoss/UnexpectedContent to MinorFormattingLoss
            if (info->get_WarningType() == WarningType::MinorFormattingLoss)
            {
                std::cout << (String(u"Unsupported operation: ") + info->get_Description()) << std::endl;
                Warnings->Warning(info);
            }
        }

        HandleDocumentWarnings() : Warnings(MakeObject<WarningInfoCollection>())
        {
        }
    };
    //ExEnd

    void HeaderFooterBookmarksExportMode(Aspose::Words::Saving::HeaderFooterBookmarksExportMode headerFooterBookmarksExportMode)
    {
        //ExStart
        //ExFor:HeaderFooterBookmarksExportMode
        //ExFor:OutlineOptions
        //ExFor:OutlineOptions.DefaultBookmarksOutlineLevel
        //ExFor:PdfSaveOptions.HeaderFooterBookmarksExportMode
        //ExFor:PdfSaveOptions.PageMode
        //ExFor:PdfPageMode
        //ExSummary:Shows how bookmarks in headers/footers are exported to pdf.
        auto doc = MakeObject<Document>(MyDir + u"Bookmarks in headers and footers.docx");

        // You can specify how bookmarks in headers/footers are exported
        // There is a several options for this:
        // "None" - Bookmarks in headers/footers are not exported
        // "First" - Only bookmark in first header/footer of the section is exported
        // "All" - Bookmarks in all headers/footers are exported
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_HeaderFooterBookmarksExportMode(headerFooterBookmarksExportMode);
        saveOptions->get_OutlineOptions()->set_DefaultBookmarksOutlineLevel(1);
        saveOptions->set_PageMode(PdfPageMode::UseOutlines);
        doc->Save(ArtifactsDir + u"PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf", saveOptions);
        //ExEnd

    }

    void UnsupportedImageFormatWarning()
    {
        auto doc = MakeObject<Document>(MyDir + u"Corrupted image.docx");

        auto saveWarningCallback = MakeObject<ExPdfSaveOptions::SaveWarningCallback>();
        doc->set_WarningCallback(saveWarningCallback);

        doc->Save(ArtifactsDir + u"PdfSaveOption.UnsupportedImageFormatWarning.pdf", SaveFormat::Pdf);

        ASSERT_EQ(saveWarningCallback->SaveWarnings->idx_get(0)->get_Description(), u"Image can not be processed. Possibly unsupported image format.");
    }

    class SaveWarningCallback : public IWarningCallback
    {
        friend class ExPdfSaveOptions;

    public:
        void Warning(SharedPtr<WarningInfo> info) override
        {
            if (info->get_WarningType() == WarningType::MinorFormattingLoss)
            {
                std::cout << String::Format(u"{0}: {1}.", info->get_WarningType(), info->get_Description()) << std::endl;
                SaveWarnings->Warning(info);
            }
        }

        SaveWarningCallback() : SaveWarnings(MakeObject<WarningInfoCollection>())
        {
        }

    protected:
        SharedPtr<WarningInfoCollection> SaveWarnings;
    };

    void FontsScaledToMetafileSize(bool doScaleWmfFonts)
    {
        //ExStart
        //ExFor:MetafileRenderingOptions.ScaleWmfFontsToMetafileSize
        //ExSummary:Shows how to WMF fonts scaling according to metafile size on the page.
        auto doc = MakeObject<Document>(MyDir + u"WMF with text.docx");

        // There is a several options for this:
        // 'True' - Aspose.Words emulates font scaling according to metafile size on the page
        // 'False' - Aspose.Words displays the fonts as metafile is rendered to its default size
        // Use 'False' option is used only when metafile is rendered as vector graphics
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->get_MetafileRenderingOptions()->set_ScaleWmfFontsToMetafileSize(doScaleWmfFonts);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.FontsScaledToMetafileSize.pdf", saveOptions);
        //ExEnd

    }

    void AdditionalTextPositioning(bool applyAdditionalTextPositioning)
    {
        //ExStart
        //ExFor:PdfSaveOptions.AdditionalTextPositioning
        //ExSummary:Show how to write additional text positioning operators.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // This may help to overcome issues with inaccurate text positioning with some printers, even if the PDF looks fine,
        // but the file size will increase due to higher text positioning precision used
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_AdditionalTextPositioning(applyAdditionalTextPositioning);
        saveOptions->set_TextCompression(PdfTextCompression::None);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
        //ExEnd

    }

    void SaveAsPdfBookFold(bool doRenderTextAsBookfold)
    {
        //ExStart
        //ExFor:PdfSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to save a document to the PDF format in the form of a book fold.
        // Open a document with multiple paragraphs
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // Configure both page setup and PdfSaveOptions to create a book fold
        for (auto s : System::IterateOver<Section>(doc->get_Sections()))
        {
            s->get_PageSetup()->set_MultiplePages(MultiplePagesType::BookFoldPrinting);
        }

        auto options = MakeObject<PdfSaveOptions>();
        options->set_UseBookFoldPrintingSettings(doRenderTextAsBookfold);

        // Once we print this document, we can turn it into a booklet by stacking the pages
        // in the order they come out of the printer and then folding down the middle
        doc->Save(ArtifactsDir + u"PdfSaveOptions.SaveAsPdfBookFold.pdf", options);
        //ExEnd

    }

    void ZoomBehaviour()
    {
        //ExStart
        //ExFor:PdfSaveOptions.ZoomBehavior
        //ExFor:PdfSaveOptions.ZoomFactor
        //ExFor:PdfZoomBehavior
        //ExSummary:Shows how to set the default zooming of an output PDF to 1/4 of default size.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<PdfSaveOptions>();
        options->set_ZoomBehavior(PdfZoomBehavior::ZoomFactor);
        options->set_ZoomFactor(25);

        // Upon opening the .pdf with a viewer such as Adobe Acrobat Pro, the zoom level will be at 25% by default,
        // with thumbnails for each page to the left
        doc->Save(ArtifactsDir + u"PdfSaveOptions.ZoomBehaviour.pdf", options);
        //ExEnd

    }

    void PageMode(PdfPageMode pageMode)
    {
        //ExStart
        //ExFor:PdfSaveOptions.PageMode
        //ExFor:PdfPageMode
        //ExSummary:Shows how to set instructions for some PDF readers to follow when opening an output document.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto options = MakeObject<PdfSaveOptions>();
        options->set_PageMode(pageMode);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.PageMode.pdf", options);
        //ExEnd

        String docLocaleName = MakeObject<System::Globalization::CultureInfo>(doc->get_Styles()->get_DefaultFont()->get_LocaleId())->get_Name();

        switch (pageMode)
        {
        case PdfPageMode::FullScreen:
            TestUtil::FileContainsString(String::Format(u"<</Type /Catalog/Pages 3 0 R/PageMode /FullScreen/Lang({0})>>\r\n", docLocaleName),
                                         ArtifactsDir + u"PdfSaveOptions.PageMode.pdf");
            break;

        case PdfPageMode::UseThumbs:
            TestUtil::FileContainsString(String::Format(u"<</Type /Catalog/Pages 3 0 R/PageMode /UseThumbs/Lang({0})>>", docLocaleName),
                                         ArtifactsDir + u"PdfSaveOptions.PageMode.pdf");
            break;

        case PdfPageMode::UseOC:
            TestUtil::FileContainsString(String::Format(u"<</Type /Catalog/Pages 3 0 R/PageMode /UseOC/Lang({0})>>\r\n", docLocaleName),
                                         ArtifactsDir + u"PdfSaveOptions.PageMode.pdf");
            break;

        case PdfPageMode::UseOutlines:
        case PdfPageMode::UseNone:
            TestUtil::FileContainsString(String::Format(u"<</Type /Catalog/Pages 3 0 R/Lang({0})>>\r\n", docLocaleName),
                                         ArtifactsDir + u"PdfSaveOptions.PageMode.pdf");
            break;
        }
    }

    void NoteHyperlinks(bool doCreateHyperlinks)
    {
        //ExStart
        //ExFor:PdfSaveOptions.CreateNoteHyperlinks
        //ExSummary:Shows how to make footnotes and endnotes work like hyperlinks.
        // Open a document with footnotes/endnotes
        auto doc = MakeObject<Document>(MyDir + u"Footnotes and endnotes.docx");

        // Creating a PdfSaveOptions instance with this flag set will convert footnote/endnote number symbols in the text
        // into hyperlinks pointing to the footnotes, and the actual footnotes/endnotes at the end of pages into links to their
        // referenced body text
        auto options = MakeObject<PdfSaveOptions>();
        options->set_CreateNoteHyperlinks(doCreateHyperlinks);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf", options);
        //ExEnd

        if (doCreateHyperlinks)
        {
            TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect [157.80099487 720.90106201 159.35600281 733.55004883]/BS <</Type/Border/S/S/W "
                                         u"0>>/Dest[4 0 R /XYZ 85 677 0]>>",
                                         ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect [202.16900635 720.90106201 206.06201172 733.55004883]/BS <</Type/Border/S/S/W "
                                         u"0>>/Dest[4 0 R /XYZ 85 79 0]>>",
                                         ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect [212.23199463 699.2510376 215.34199524 711.90002441]/BS <</Type/Border/S/S/W "
                                         u"0>>/Dest[4 0 R /XYZ 85 654 0]>>",
                                         ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect [258.15499878 699.2510376 262.04800415 711.90002441]/BS <</Type/Border/S/S/W "
                                         u"0>>/Dest[4 0 R /XYZ 85 68 0]>>",
                                         ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(
                u"<</Type /Annot/Subtype /Link/Rect [85.05000305 68.19905853 88.66500092 79.69805908]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 202 733 0]>>",
                ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(
                u"<</Type /Annot/Subtype /Link/Rect [85.05000305 56.70005798 88.66500092 68.19905853]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 258 711 0]>>",
                ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect [85.05000305 666.10205078 86.4940033 677.60107422]/BS <</Type/Border/S/S/W "
                                         u"0>>/Dest[4 0 R /XYZ 157 733 0]>>",
                                         ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
            TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect [85.05000305 643.10406494 87.93800354 654.60308838]/BS <</Type/Border/S/S/W "
                                         u"0>>/Dest[4 0 R /XYZ 212 711 0]>>",
                                         ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf");
        }
        else
        {
            if (!IsRunningOnMono())
            {

                EXPECT_FATAL_FAILURE(TestUtil::FileContainsString(u"<</Type /Annot/Subtype /Link/Rect", ArtifactsDir + u"PdfSaveOptions.NoteHyperlinks.pdf"),
                    "not found in the provided source");
            }
        }
    }

    void CustomPropertiesExport(PdfCustomPropertiesExport pdfCustomPropertiesExportMode)
    {
        //ExStart
        //ExFor:PdfCustomPropertiesExport
        //ExFor:PdfSaveOptions.CustomPropertiesExport
        //ExSummary:Shows how to export custom properties while saving to .pdf.
        auto doc = MakeObject<Document>();

        // Add a custom document property that does not use the name of some built in properties
        doc->get_CustomDocumentProperties()->Add(u"Company", String(u"My value"));

        // Configure the PdfSaveOptions like this will display the properties
        // in the "Document Properties" menu of Adobe Acrobat Pro
        auto options = MakeObject<PdfSaveOptions>();
        options->set_CustomPropertiesExport(pdfCustomPropertiesExportMode);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.CustomPropertiesExport.pdf", options);
        //ExEnd

        switch (pdfCustomPropertiesExportMode)
        {
        case PdfCustomPropertiesExport::None:
            if (!IsRunningOnMono())
            {

                EXPECT_FATAL_FAILURE(TestUtil::FileContainsString(u"Company", ArtifactsDir + u"PdfSaveOptions.CustomPropertiesExport.pdf"),
                        "not found in the provided source");

                    EXPECT_FATAL_FAILURE(TestUtil::FileContainsString(u"<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>", ArtifactsDir + u"PdfSaveOptions.CustomPropertiesExport.pdf"),
                        "not found in the provided source");
            }
            break;

        case PdfCustomPropertiesExport::Standard:
            TestUtil::FileContainsString(doc->get_CustomDocumentProperties()->idx_get(0)->get_Name(),
                                         ArtifactsDir + u"PdfSaveOptions.CustomPropertiesExport.pdf");
            break;

        case PdfCustomPropertiesExport::Metadata:
            TestUtil::FileContainsString(u"<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
                                         ArtifactsDir + u"PdfSaveOptions.CustomPropertiesExport.pdf");
            break;
        }
    }

    void DrawingMLEffects(DmlEffectsRenderingMode effectsRenderingMode)
    {
        //ExStart
        //ExFor:DmlRenderingMode
        //ExFor:DmlEffectsRenderingMode
        //ExFor:PdfSaveOptions.DmlEffectsRenderingMode
        //ExFor:SaveOptions.DmlEffectsRenderingMode
        //ExFor:SaveOptions.DmlRenderingMode
        //ExSummary:Shows how to configure DrawingML rendering quality with PdfSaveOptions.
        auto doc = MakeObject<Document>(MyDir + u"DrawingML shape effects.docx");

        auto options = MakeObject<PdfSaveOptions>();
        options->set_DmlEffectsRenderingMode(effectsRenderingMode);

        ASSERT_EQ(DmlRenderingMode::DrawingML, options->get_DmlRenderingMode());

        doc->Save(ArtifactsDir + u"PdfSaveOptions.DrawingMLEffects.pdf", options);
        //ExEnd

    }

    void DrawingMLFallback(DmlRenderingMode dmlRenderingMode)
    {
        //ExStart
        //ExFor:DmlRenderingMode
        //ExFor:SaveOptions.DmlRenderingMode
        //ExSummary:Shows how to render fallback shapes when saving to Pdf.
        auto doc = MakeObject<Document>(MyDir + u"DrawingML shape fallbacks.docx");

        auto options = MakeObject<PdfSaveOptions>();
        options->set_DmlRenderingMode(dmlRenderingMode);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.DrawingMLFallback.pdf", options);
        //ExEnd

        switch (dmlRenderingMode)
        {
        case DmlRenderingMode::DrawingML:
            TestUtil::FileContainsString(u"<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 "
                                         u"R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                                         ArtifactsDir + u"PdfSaveOptions.DrawingMLFallback.pdf");
            break;

        case DmlRenderingMode::Fallback:
            TestUtil::FileContainsString(u"4 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABC "
                                         u"12 0 R>>/ExtGState<</GS1 9 0 R/GS2 10 0 R>>>>/Group ",
                                         ArtifactsDir + u"PdfSaveOptions.DrawingMLFallback.pdf");
            break;
        }
    }

    void ExportDocumentStructure(bool doExportStructure)
    {
        //ExStart
        //ExFor:PdfSaveOptions.ExportDocumentStructure
        //ExSummary:Shows how to convert a .docx to .pdf while preserving the document structure.
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // Create a PdfSaveOptions object and configure it to preserve the logical structure that's in the input document
        // The file size will be increased, and the structure will be visible in the "Content" navigation pane of Adobe Acrobat Pro
        auto options = MakeObject<PdfSaveOptions>();
        options->set_ExportDocumentStructure(doExportStructure);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.ExportDocumentStructure.pdf", options);
        //ExEnd

        if (doExportStructure)
        {
            TestUtil::FileContainsString(
                String(u"4 0 obj\r\n") +
                    u"<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABC 12 0 R>>/ExtGState<</GS1 9 0 "
                    u"R/GS2 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>/StructParents 0/Tabs /S>>",
                ArtifactsDir + u"PdfSaveOptions.ExportDocumentStructure.pdf");
        }
        else
        {
            TestUtil::FileContainsString(String(u"4 0 obj\r\n") +
                                             u"<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 "
                                             u"R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                                         ArtifactsDir + u"PdfSaveOptions.ExportDocumentStructure.pdf");
        }
    }

    void PreblendImages(bool doPreblendImages)
    {
        //ExStart
        //ExFor:PdfSaveOptions.PreblendImages
        //ExSummary:Shows how to preblend images with transparent backgrounds.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<System::Drawing::Image> img = System::Drawing::Image::FromFile(ImageDir + u"Transparent background logo.png");
        builder->InsertImage(img);

        // Setting this flag in a SaveOptions object may change the quality and size of the output .pdf
        // because of the way some images are rendered
        auto options = MakeObject<PdfSaveOptions>();
        options->set_PreblendImages(doPreblendImages);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.PreblendImages.pdf", options);
        //ExEnd
    }

    void InterpolateImages()
    {
        //ExStart
        //ExFor:PdfSaveOptions.InterpolateImages
        //ExSummary:Shows how to improve the quality of an image in the rendered documents.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<System::Drawing::Image> img = System::Drawing::Image::FromFile(ImageDir + u"Transparent background logo.png");
        builder->InsertImage(img);

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_InterpolateImages(true);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.InterpolateImages.pdf", saveOptions);
        //ExEnd
    }

    void Dml3DEffectsRenderingModeTest()
    {
        auto doc = MakeObject<Document>(MyDir + u"DrawingML shape 3D effects.docx");

        auto warningCallback = MakeObject<ExPdfSaveOptions::RenderCallback>();
        doc->set_WarningCallback(warningCallback);

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_Dml3DEffectsRenderingMode(Dml3DEffectsRenderingMode::Advanced);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.Dml3DEffectsRenderingModeTest.pdf", saveOptions);

        ASSERT_EQ(43, warningCallback->get_Count());
    }

    class RenderCallback : public IWarningCallback
    {
    public:
        int get_Count()
        {
            return mWarnings->get_Count();
        }

        void Warning(SharedPtr<WarningInfo> info) override
        {
            std::cout << String::Format(u"{0}: {1}.", info->get_WarningType(), info->get_Description()) << std::endl;
            mWarnings->Add(info);
        }

        SharedPtr<WarningInfo> idx_get(int i)
        {
            return mWarnings->idx_get(i);
        }

        /// <summary>
        /// Clears warning collection.
        /// </summary>
        void Clear()
        {
            mWarnings->Clear();
        }

        /// <summary>
        /// Returns true if a warning with the specified properties has been generated.
        /// </summary>
        bool Contains(WarningSource source, WarningType type, String description)
        {

            // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
            std::function<bool(SharedPtr<WarningInfo> warning)> _local_func_3 = [&source, &type, &description](SharedPtr<WarningInfo> warning)
            {
                return warning->get_Source() == source && warning->get_WarningType() == type && warning->get_Description() == description;
            };

            return mWarnings->LINQ_Any(static_cast<System::Func<SharedPtr<WarningInfo>, bool>>(_local_func_3));
        }

        RenderCallback() : mWarnings(MakeObject<System::Collections::Generic::List<SharedPtr<WarningInfo>>>())
        {
        }

    private:
        SharedPtr<System::Collections::Generic::List<SharedPtr<WarningInfo>>> mWarnings;
    };

    void PdfDigitalSignature()
    {
        //ExStart
        //ExFor:PdfDigitalSignatureDetails
        //ExFor:PdfDigitalSignatureDetails.#ctor(CertificateHolder, String, String, DateTime)
        //ExFor:PdfDigitalSignatureDetails.HashAlgorithm
        //ExFor:PdfDigitalSignatureDetails.Location
        //ExFor:PdfDigitalSignatureDetails.Reason
        //ExFor:PdfDigitalSignatureDetails.SignatureDate
        //ExFor:PdfDigitalSignatureHashAlgorithm
        //ExFor:PdfSaveOptions.DigitalSignatureDetails
        //ExSummary:Shows how to sign a generated PDF using Aspose.Words.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Signed PDF contents.");

        // Load the certificate from disk
        // The other constructor overloads can be used to load certificates from different locations
        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        // Pass the certificate and details to the save options class to sign with
        auto options = MakeObject<PdfSaveOptions>();
        System::DateTime signingTime = System::DateTime::get_Now();
        options->set_DigitalSignatureDetails(MakeObject<PdfDigitalSignatureDetails>(certificateHolder, u"Test Signing", u"Aspose Office", signingTime));

        // We can use this attribute to set a different hash algorithm
        options->get_DigitalSignatureDetails()->set_HashAlgorithm(PdfDigitalSignatureHashAlgorithm::Sha256);

        ASSERT_EQ(u"Test Signing", options->get_DigitalSignatureDetails()->get_Reason());
        ASSERT_EQ(u"Aspose Office", options->get_DigitalSignatureDetails()->get_Location());
        ASSERT_EQ(signingTime.ToUniversalTime(), options->get_DigitalSignatureDetails()->get_SignatureDate());

        doc->Save(ArtifactsDir + u"PdfSaveOptions.PdfDigitalSignature.pdf", options);
        //ExEnd

        TestUtil::FileContainsString(String(u"6 0 obj\r\n") +
                                         u"<</Type /Annot/Subtype /Widget/FT /Sig/DR <<>>/F 132/Rect [0 0 0 0]/V 7 0 R/P 4 0 "
                                         u"R/"
                                         u"T("
                                         u"þÿ\u0000A\u0000s\u0000p\u0000o\u0000s\u0000e\u0000D\u0000i\u0000g\u0000i\u0000t\u0000a\u0000l\u0000S\u0000i\u0000g"
                                         u"\u0000n\u0000a\u0000t\u0000u\u0000r\u0000e)/AP <</N 8 0 R>>>>",
                                     ArtifactsDir + u"PdfSaveOptions.PdfDigitalSignature.pdf");

        ASSERT_FALSE(FileFormatUtil::DetectFileFormat(ArtifactsDir + u"PdfSaveOptions.PdfDigitalSignature.pdf")->get_HasDigitalSignature());
    }

    void PdfDigitalSignatureTimestamp()
    {
        //ExStart
        //ExFor:PdfDigitalSignatureDetails.TimestampSettings
        //ExFor:PdfDigitalSignatureTimestampSettings
        //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String)
        //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String,TimeSpan)
        //ExFor:PdfDigitalSignatureTimestampSettings.Password
        //ExFor:PdfDigitalSignatureTimestampSettings.ServerUrl
        //ExFor:PdfDigitalSignatureTimestampSettings.Timeout
        //ExFor:PdfDigitalSignatureTimestampSettings.UserName
        //ExSummary:Shows how to sign a generated PDF and timestamp it using Aspose.Words.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Signed PDF contents.");

        // Create a digital signature for the document that we will save
        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");
        auto options = MakeObject<PdfSaveOptions>();
        options->set_DigitalSignatureDetails(
            MakeObject<PdfDigitalSignatureDetails>(certificateHolder, u"Test Signing", u"Aspose Office", System::DateTime::get_Now()));

        // We can set a verified timestamp for our signature as well, with a valid timestamp authority
        options->get_DigitalSignatureDetails()->set_TimestampSettings(
            MakeObject<PdfDigitalSignatureTimestampSettings>(u"https://freetsa.org/tsr", u"JohnDoe", u"MyPassword"));

        // The default lifespan of the timestamp is 100 seconds
        ASPOSE_ASSERT_EQ(100.0, options->get_DigitalSignatureDetails()->get_TimestampSettings()->get_Timeout().get_TotalSeconds());

        // We can set our own timeout period via the constructor
        options->get_DigitalSignatureDetails()->set_TimestampSettings(
            MakeObject<PdfDigitalSignatureTimestampSettings>(u"https://freetsa.org/tsr", u"JohnDoe", u"MyPassword", System::TimeSpan::FromMinutes(30)));

        ASPOSE_ASSERT_EQ(1800.0, options->get_DigitalSignatureDetails()->get_TimestampSettings()->get_Timeout().get_TotalSeconds());
        ASSERT_EQ(u"https://freetsa.org/tsr", options->get_DigitalSignatureDetails()->get_TimestampSettings()->get_ServerUrl());
        ASSERT_EQ(u"JohnDoe", options->get_DigitalSignatureDetails()->get_TimestampSettings()->get_UserName());
        ASSERT_EQ(u"MyPassword", options->get_DigitalSignatureDetails()->get_TimestampSettings()->get_Password());

        doc->Save(ArtifactsDir + u"PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf", options);
        //ExEnd

        ASSERT_FALSE(FileFormatUtil::DetectFileFormat(ArtifactsDir + u"PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf")->get_HasDigitalSignature());
        TestUtil::FileContainsString(String(u"6 0 obj\r\n") +
                                         u"<</Type /Annot/Subtype /Widget/FT /Sig/DR <<>>/F 132/Rect [0 0 0 0]/V 7 0 R/P 4 0 "
                                         u"R/"
                                         u"T("
                                         u"þÿ\u0000A\u0000s\u0000p\u0000o\u0000s\u0000e\u0000D\u0000i\u0000g\u0000i\u0000t\u0000a\u0000l\u0000S\u0000i\u0000g"
                                         u"\u0000n\u0000a\u0000t\u0000u\u0000r\u0000e)/AP <</N 8 0 R>>>>",
                                     ArtifactsDir + u"PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");
    }

    void RenderMetafile(EmfPlusDualRenderingMode renderingMode)
    {
        //ExStart
        //ExFor:EmfPlusDualRenderingMode
        //ExFor:MetafileRenderingOptions.EmfPlusDualRenderingMode
        //ExFor:MetafileRenderingOptions.UseEmfEmbeddedToWmf
        //ExSummary:Shows how to adjust EMF (Enhanced Windows Metafile) rendering options when saving to PDF.
        auto doc = MakeObject<Document>(MyDir + u"EMF.docx");

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->get_MetafileRenderingOptions()->set_EmfPlusDualRenderingMode(renderingMode);
        saveOptions->get_MetafileRenderingOptions()->set_UseEmfEmbeddedToWmf(true);

        doc->Save(ArtifactsDir + u"PdfSaveOptions.RenderMetafile.pdf", saveOptions);
        //ExEnd

    }
};

} // namespace ApiExamples
