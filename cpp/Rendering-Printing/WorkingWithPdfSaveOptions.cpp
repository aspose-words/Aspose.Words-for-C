#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/HeaderFooterBookmarksExportMode.h>
#include <Aspose.Words.Cpp/Model/Saving/MetafileRenderingOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OutlineOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
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
}

void WorkingWithPdfSaveOptions()
{
    std::cout << "WorkingWithPdfSaveOptions example started." << std::endl;
    // ExStart:SetTrueTypeFontsFolder
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();
    EscapeUriInPdf(inputDataDir, outputDataDir);
    ExportHeaderFooterBookmarks(inputDataDir, outputDataDir);
    ScaleWmfFontsToMetafileSize(inputDataDir, outputDataDir);
    AdditionalTextPositioning(inputDataDir, outputDataDir);
    std::cout << "WorkingWithPdfSaveOptions example finished." << std::endl << std::endl;
}