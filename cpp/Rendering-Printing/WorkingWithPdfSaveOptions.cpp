#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Saving/HeaderFooterBookmarksExportMode.h>
#include <Model/Saving/MetafileRenderingOptions.h>
#include <Model/Saving/OutlineOptions.h>
#include <Model/Saving/PdfSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void EscapeUriInPdf(System::String const &dataDir)
    {
        // ExStart:EscapeUriInPdf
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"EscapeUri.docx");

        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_EscapeUri(false);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithPdfSaveOptions.EscapeUriInPdf.pdf");
        doc->Save(outputPath, options);
        // ExEnd:EscapeUriInPdf
        std::cout << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ExportHeaderFooterBookmarks(System::String const &dataDir)
    {
        // ExStart:ExportHeaderFooterBookmarks
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.docx");


        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->get_OutlineOptions()->set_DefaultBookmarksOutlineLevel(1);
        options->set_HeaderFooterBookmarksExportMode(HeaderFooterBookmarksExportMode::First);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithPdfSaveOptions.ExportHeaderFooterBookmarks.pdf");
        doc->Save(outputPath, options);
        // ExEnd:ExportHeaderFooterBookmarks
        std::cout << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ScaleWmfFontsToMetafileSize(System::String const &dataDir)
    {
        // ExStart:ScaleWmfFontsToMetafileSize
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"MetafileRendering.docx");

        System::SharedPtr<MetafileRenderingOptions> metafileRenderingOptions = System::MakeObject<MetafileRenderingOptions>();
        metafileRenderingOptions->set_ScaleWmfFontsToMetafileSize(false);

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap. 
        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_MetafileRenderingOptions(metafileRenderingOptions);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithPdfSaveOptions.ScaleWmfFontsToMetafileSize.pdf");
        doc->Save(outputPath, options);
        // ExEnd:ScaleWmfFontsToMetafileSize
        std::cout << "Fonts as metafile are rendered to its default size in PDF." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkingWithPdfSaveOptions()
{
    std::cout << "WorkingWithPdfSaveOptions example started." << std::endl;
    // ExStart:SetTrueTypeFontsFolder
    // The path to the documents directory.
    System::String dataDir = GetDataDir_RenderingAndPrinting();
    EscapeUriInPdf(dataDir);
    ExportHeaderFooterBookmarks(dataDir);
    ScaleWmfFontsToMetafileSize(dataDir);
    std::cout << "WorkingWithPdfSaveOptions example finished." << std::endl << std::endl;
}