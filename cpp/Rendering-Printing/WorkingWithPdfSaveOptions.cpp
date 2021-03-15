#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/HeaderFooterBookmarksExportMode.h>
#include <Aspose.Words.Cpp/Model/Saving/MetafileRenderingOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OutlineOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/DownsampleOptions.h>
#include <Aspose.Words.Cpp/Model/Properties/CustomDocumentProperties.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
#if 0 //Deprecated
    void EscapeUriInPdf(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:EscapeUriInPdf
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"EscapeUri.docx");

        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_EscapeUri(false);

        System::String outputPath = outputDataDir + u"WorkingWithPdfSaveOptions.EscapeUriInPdf.pdf";
        doc->Save(outputPath, options);
        // ExEnd:EscapeUriInPdf
        std::cout << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
#endif 

    void ExportHeaderFooterBookmarks(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:ExportHeaderFooterBookmarks
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.docx");


        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->get_OutlineOptions()->set_DefaultBookmarksOutlineLevel(1);
        options->set_HeaderFooterBookmarksExportMode(HeaderFooterBookmarksExportMode::First);

        System::String outputPath = outputDataDir + u"WorkingWithPdfSaveOptions.ExportHeaderFooterBookmarks.pdf";
        doc->Save(outputPath, options);
        // ExEnd:ExportHeaderFooterBookmarks
        std::cout << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ScaleWmfFontsToMetafileSize(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:ScaleWmfFontsToMetafileSize
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"MetafileRendering.docx");

        System::SharedPtr<MetafileRenderingOptions> metafileRenderingOptions = System::MakeObject<MetafileRenderingOptions>();
        metafileRenderingOptions->set_ScaleWmfFontsToMetafileSize(false);

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap. 
        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_MetafileRenderingOptions(metafileRenderingOptions);

        System::String outputPath = outputDataDir + u"WorkingWithPdfSaveOptions.ScaleWmfFontsToMetafileSize.pdf";
        doc->Save(outputPath, options);
        // ExEnd:ScaleWmfFontsToMetafileSize
        std::cout << "Fonts as metafile are rendered to its default size in PDF." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void AdditionalTextPositioning(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:AdditionalTextPositioning
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.docx");

        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_AdditionalTextPositioning(true);

        System::String outputPath = outputDataDir + u"WorkingWithPdfSaveOptions.AdditionalTextPositioning.pdf";
        doc->Save(outputPath, options);
        // ExEnd:AdditionalTextPositioning
        std::cout << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ConversionToPDF17(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:ConversionToPDF17
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.docx");

        // Provide PDFSaveOption compliance to PDF17
        // or just convert without SaveOptions
        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_Compliance(PdfCompliance::Pdf17);

        System::String outputPath = outputDataDir + u"WorkingWithPdfSaveOptions.ConversionToPDF17.pdf";
        doc->Save(outputPath, options);
        // ExEnd:ConversionToPDF17
        std::cout << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void UpdateIfLastPrinted(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:UpdateIfLastPrinted
        // Open a document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.docx");

        System::SharedPtr<SaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_UpdateLastPrintedProperty(false);

        System::String outputPath = outputDataDir + u"WorkingWithPdfSaveOptions.UpdateIfLastPrinted.pdf";
        doc->Save(outputPath, options);
        // ExEnd:UpdateIfLastPrinted
        std::cout << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void PdfImageComppression(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:PdfImageComppression
        // Open a document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.docx");

        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_ImageCompression(PdfImageCompression::Jpeg);
        options->set_PreserveFormFields(true);

        System::String outputPath = outputDataDir + u"SaveOptions.PdfImageCompression.pdf";
        doc->Save(outputPath, options);

        System::SharedPtr<PdfSaveOptions> optionsA1B = System::MakeObject<PdfSaveOptions>();
        optionsA1B->set_Compliance(PdfCompliance::PdfA1b);
        optionsA1B->set_ImageCompression(PdfImageCompression::Jpeg);
        optionsA1B->set_PreserveFormFields(true);
        
        // Use JPEG compression at 50% quality to reduce file size
        optionsA1B->set_JpegQuality(100);
        optionsA1B->set_ImageColorSpaceExportMode(PdfImageColorSpaceExportMode::SimpleCmyk);

        System::String outputPathA1B = outputDataDir + u"SaveOptions.PdfImageComppression_PDF_A_1_B.pdf";
        doc->Save(outputPathA1B, options);
        // ExEnd:PdfImageComppression
    }

    void ExportDocumentStructure(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:ExportDocumentStructure
        // Open a document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Paragraphs.docx");

        // Create a PdfSaveOptions object and configure it to preserve the logical structure that's in the input document
        // The file size will be increased and the structure will be visible in the "Content" navigation pane
        // of Adobe Acrobat Pro, while editing the .pdf
        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_ExportDocumentStructure(true);

        System::String outputPath = outputDataDir + u"PdfSaveOptions.ExportDocumentStructure.pdf";
        doc->Save(outputPath, options);
        // ExEnd:ExportDocumentStructure
    }

    void CustomPropertiesExport(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:CustomPropertiesExport
        // Open a document
        System::SharedPtr<Document> doc = System::MakeObject<Document>();

        // Add a custom document property that doesn't use the name of some built in properties
        doc->get_CustomDocumentProperties()->Add(u"Company", u"My value");

        // Configure the PdfSaveOptions like this will display the properties
        // in the "Document Properties" menu of Adobe Acrobat Pro
        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_CustomPropertiesExport(PdfCustomPropertiesExport::Standard);

        System::String outputPath = outputDataDir + u"PdfSaveOptions.CustomPropertiesExport.pdf";
        doc->Save(outputPath, options);
        // ExEnd:CustomPropertiesExport
    }

    void SaveToPdfWithOutline(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:SaveToPdfWithOutline
        // Open a document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->get_OutlineOptions()->set_HeadingsOutlineLevels(3);
        options->get_OutlineOptions()->set_ExpandedOutlineLevels(1);

        System::String outputPath = outputDataDir + u"PdfSaveOptions.SaveToPdfWithOutline.pdf";
        doc->Save(outputPath, options);
        // ExEnd:SaveToPdfWithOutline
    }

    void DownsamplingImages(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:DownsamplingImages
        // Open a document that contains images 
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

        // If we want to convert the document to .pdf, we can use a SaveOptions implementation to customize the saving process
        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();

        // We can set the output resolution to a different value
        // The first two images in the input document will be affected by this
        options->get_DownsampleOptions()->set_Resolution(36);

        // We can set a minimum threshold for downsampling 
        // This value will prevent the second image in the input document from being downsampled
        options->get_DownsampleOptions()->set_ResolutionThreshold(128);

        System::String outputPath = outputDataDir + u"PdfSaveOptions.DownsampleOptions.pdf";
        doc->Save(outputPath, options);
        // ExEnd:DownsamplingImages
    }

	void SetImageInterpolation(System::String const& inputPath, System::String const& outputDataDir)
	{
		// ExStart:SetImageInterpolation
        auto doc = System::MakeObject<Aspose::Words::Document>(inputPath + u"Rendering.docx");

        System::SharedPtr<PdfSaveOptions> saveOptions = System::MakeObject<PdfSaveOptions>();
		saveOptions->set_InterpolateImages(true);

        doc->Save(outputDataDir + u"WorkingWithPdfSaveOptions.InterpolateImages.pdf", saveOptions);
		// ExEnd:SetImageInterpolation
	}
}

void WorkingWithPdfSaveOptions()
{
    std::cout << "WorkingWithPdfSaveOptions example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    // Deprecated 
    // EscapeUriInPdf(inputDataDir, outputDataDir);

    ExportHeaderFooterBookmarks(inputDataDir, outputDataDir);
    ScaleWmfFontsToMetafileSize(inputDataDir, outputDataDir);
    AdditionalTextPositioning(inputDataDir, outputDataDir);
    ConversionToPDF17(inputDataDir, outputDataDir);
    UpdateIfLastPrinted(inputDataDir, outputDataDir);
    PdfImageComppression(inputDataDir, outputDataDir);
    // ExportDocumentStructure(inputDataDir, outputDataDir);
    CustomPropertiesExport(inputDataDir, outputDataDir);
    SaveToPdfWithOutline(inputDataDir, outputDataDir);
    DownsamplingImages(inputDataDir, outputDataDir);
    SetImageInterpolation(inputDataDir, outputDataDir);
    std::cout << "WorkingWithPdfSaveOptions example finished." << std::endl << std::endl;
}