#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/DigitalSignatures/CertificateHolder.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/IWarningCallback.h>
#include <Aspose.Words.Cpp/Properties/CustomDocumentProperties.h>
#include <Aspose.Words.Cpp/Properties/DocumentProperty.h>
#include <Aspose.Words.Cpp/Saving/Dml3DEffectsRenderingMode.h>
#include <Aspose.Words.Cpp/Saving/DownsampleOptions.h>
#include <Aspose.Words.Cpp/Saving/HeaderFooterBookmarksExportMode.h>
#include <Aspose.Words.Cpp/Saving/MetafileRenderingMode.h>
#include <Aspose.Words.Cpp/Saving/MetafileRenderingOptions.h>
#include <Aspose.Words.Cpp/Saving/OutlineOptions.h>
#include <Aspose.Words.Cpp/Saving/PdfCompliance.h>
#include <Aspose.Words.Cpp/Saving/PdfCustomPropertiesExport.h>
#include <Aspose.Words.Cpp/Saving/PdfDigitalSignatureDetails.h>
#include <Aspose.Words.Cpp/Saving/PdfFontEmbeddingMode.h>
#include <Aspose.Words.Cpp/Saving/PdfImageCompression.h>
#include <Aspose.Words.Cpp/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/WarningInfo.h>
#include <Aspose.Words.Cpp/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/WarningType.h>
#include <system/date_time.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Saving;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options {

class WorkingWithPdfSaveOptions : public DocsExamplesBase
{
public:
    class HandleDocumentWarnings : public IWarningCallback
    {
    public:
        SharedPtr<WarningInfoCollection> mWarnings;

        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during
        /// document load and/or document save.
        /// </summary>
        void Warning(SharedPtr<WarningInfo> info) override
        {
            // For now type of warnings about unsupported metafile records changed
            // from DataLoss/UnexpectedContent to MinorFormattingLoss.
            if (info->get_WarningType() == WarningType::MinorFormattingLoss)
            {
                std::cout << (String(u"Unsupported operation: ") + info->get_Description()) << std::endl;
                mWarnings->Warning(info);
            }
        }

        HandleDocumentWarnings() : mWarnings(MakeObject<WarningInfoCollection>())
        {
        }
    };

public:
    void DisplayDocTitleInWindowTitlebar()
    {
        //ExStart:DisplayDocTitleInWindowTitlebar
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_DisplayDocTitle(true);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.DisplayDocTitleInWindowTitlebar.pdf", saveOptions);
        //ExEnd:DisplayDocTitleInWindowTitlebar
    }

    void PdfRenderWarnings()
    {
        auto doc = MakeObject<Document>(MyDir + u"WMF with image.docx");

        auto metafileRenderingOptions = MakeObject<MetafileRenderingOptions>();
        metafileRenderingOptions->set_EmulateRasterOperations(false);
        metafileRenderingOptions->set_RenderingMode(MetafileRenderingMode::VectorWithFallback);

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_MetafileRenderingOptions(metafileRenderingOptions);

        // If Aspose.Words cannot correctly render some of the metafile records
        // to vector graphics then Aspose.Words renders this metafile to a bitmap.
        auto callback = MakeObject<WorkingWithPdfSaveOptions::HandleDocumentWarnings>();
        doc->set_WarningCallback(callback);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.PdfRenderWarnings.pdf", saveOptions);

        // While the file saves successfully, rendering warnings that occurred during saving are collected here.
        for (const auto& warningInfo : callback->mWarnings)
        {
            std::cout << warningInfo->get_Description() << std::endl;
        }
    }

    void DigitallySignedPdfUsingCertificateHolder()
    {
        //ExStart:DigitallySignedPdfUsingCertificateHolder
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Test Signed PDF.");

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_DigitalSignatureDetails(MakeObject<PdfDigitalSignatureDetails>(CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw"), u"reason",
                                                                                        u"location", System::DateTime::get_Now()));

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.DigitallySignedPdfUsingCertificateHolder.pdf", saveOptions);
        //ExEnd:DigitallySignedPdfUsingCertificateHolder
    }

    void EmbeddedAllFonts()
    {
        //ExStart:EmbeddAllFonts
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // The output PDF will be embedded with all fonts found in the document.
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_EmbedFullFonts(true);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.EmbeddedFontsInPdf.pdf", saveOptions);
        //ExEnd:EmbeddAllFonts
    }

    void EmbeddedSubsetFonts()
    {
        //ExStart:EmbeddSubsetFonts
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // The output PDF will contain subsets of the fonts in the document.
        // Only the glyphs used in the document are included in the PDF fonts.
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_EmbedFullFonts(false);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.EmbeddSubsetFonts.pdf", saveOptions);
        //ExEnd:EmbeddSubsetFonts
    }

    void DisableEmbedWindowsFonts()
    {
        //ExStart:DisableEmbedWindowsFonts
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // The output PDF will be saved without embedding standard windows fonts.
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_FontEmbeddingMode(PdfFontEmbeddingMode::EmbedNone);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.DisableEmbedWindowsFonts.pdf", saveOptions);
        //ExEnd:DisableEmbedWindowsFonts
    }

    void SkipEmbeddedArialAndTimesRomanFonts()
    {
        //ExStart:SkipEmbeddedArialAndTimesRomanFonts
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_FontEmbeddingMode(PdfFontEmbeddingMode::EmbedAll);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.SkipEmbeddedArialAndTimesRomanFonts.pdf", saveOptions);
        //ExEnd:SkipEmbeddedArialAndTimesRomanFonts
    }

    void AvoidEmbeddingCoreFonts()
    {
        //ExStart:AvoidEmbeddingCoreFonts
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_UseCoreFonts(true);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.AvoidEmbeddingCoreFonts.pdf", saveOptions);
        //ExEnd:AvoidEmbeddingCoreFonts
    }

    void ExportHeaderFooterBookmarks()
    {
        //ExStart:ExportHeaderFooterBookmarks
        auto doc = MakeObject<Document>(MyDir + u"Bookmarks in headers and footers.docx");

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->get_OutlineOptions()->set_DefaultBookmarksOutlineLevel(1);
        saveOptions->set_HeaderFooterBookmarksExportMode(HeaderFooterBookmarksExportMode::First);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.ExportHeaderFooterBookmarks.pdf", saveOptions);
        //ExEnd:ExportHeaderFooterBookmarks
    }

    void ScaleWmfFontsToMetafileSize()
    {
        //ExStart:ScaleWmfFontsToMetafileSize
        auto doc = MakeObject<Document>(MyDir + u"WMF with text.docx");

        auto metafileRenderingOptions = MakeObject<MetafileRenderingOptions>();
        metafileRenderingOptions->set_ScaleWmfFontsToMetafileSize(false);

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics
        // then Aspose.Words renders this metafile to a bitmap.
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_MetafileRenderingOptions(metafileRenderingOptions);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.ScaleWmfFontsToMetafileSize.pdf", saveOptions);
        //ExEnd:ScaleWmfFontsToMetafileSize
    }

    void AdditionalTextPositioning()
    {
        //ExStart:AdditionalTextPositioning
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_AdditionalTextPositioning(true);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
        //ExEnd:AdditionalTextPositioning
    }

    void ConversionToPdf17()
    {
        //ExStart:ConversionToPDF17
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_Compliance(PdfCompliance::Pdf17);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.ConversionToPdf17.pdf", saveOptions);
        //ExEnd:ConversionToPDF17
    }

    void DownsamplingImages()
    {
        //ExStart:DownsamplingImages
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // We can set a minimum threshold for downsampling.
        // This value will prevent the second image in the input document from being downsampled.
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->get_DownsampleOptions()->set_Resolution(36);
        saveOptions->get_DownsampleOptions()->set_ResolutionThreshold(128);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.DownsamplingImages.pdf", saveOptions);
        //ExEnd:DownsamplingImages
    }

    void SetOutlineOptions()
    {
        //ExStart:SetOutlineOptions
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->get_OutlineOptions()->set_HeadingsOutlineLevels(3);
        saveOptions->get_OutlineOptions()->set_ExpandedOutlineLevels(1);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.SetOutlineOptions.pdf", saveOptions);
        //ExEnd:SetOutlineOptions
    }

    void CustomPropertiesExport()
    {
        //ExStart:CustomPropertiesExport
        auto doc = MakeObject<Document>();
        doc->get_CustomDocumentProperties()->Add(u"Company", String(u"Aspose"));

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_CustomPropertiesExport(PdfCustomPropertiesExport::Standard);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.CustomPropertiesExport.pdf", saveOptions);
        //ExEnd:CustomPropertiesExport
    }

    void ExportDocumentStructure()
    {
        //ExStart:ExportDocumentStructure
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // The file size will be increased and the structure will be visible in the "Content" navigation pane
        // of Adobe Acrobat Pro, while editing the .pdf.
        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_ExportDocumentStructure(true);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.ExportDocumentStructure.pdf", saveOptions);
        //ExEnd:ExportDocumentStructure
    }

    void ImageCompression()
    {
        //ExStart:PdfImageCompression
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_ImageCompression(PdfImageCompression::Jpeg);
        saveOptions->set_PreserveFormFields(true);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.PdfImageCompression.pdf", saveOptions);

        auto saveOptionsA2U = MakeObject<PdfSaveOptions>();
        saveOptionsA2U->set_Compliance(PdfCompliance::PdfA2u);
        saveOptionsA2U->set_ImageCompression(PdfImageCompression::Jpeg);
        saveOptionsA2U->set_JpegQuality(100);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.PdfImageCompression_A2u.pdf", saveOptionsA2U);
        //ExEnd:PdfImageCompression
    }

    void UpdateLastPrintedProperty()
    {
        //ExStart:UpdateIfLastPrinted
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_UpdateLastPrintedProperty(true);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.UpdateIfLastPrinted.pdf", saveOptions);
        //ExEnd:UpdateIfLastPrinted
    }

    void Dml3DEffectsRendering()
    {
        //ExStart:Dml3DEffectsRendering
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_Dml3DEffectsRenderingMode(Dml3DEffectsRenderingMode::Advanced);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.Dml3DEffectsRendering.pdf", saveOptions);
        //ExEnd:Dml3DEffectsRendering
    }

    void InterpolateImages()
    {
        //ExStart:SetImageInterpolation
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PdfSaveOptions>();
        saveOptions->set_InterpolateImages(true);

        doc->Save(ArtifactsDir + u"WorkingWithPdfSaveOptions.InterpolateImages.pdf", saveOptions);
        //ExEnd:SetImageInterpolation
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
