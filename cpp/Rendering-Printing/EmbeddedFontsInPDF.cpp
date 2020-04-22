#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void EmbedAllFonts(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:EmbeddAllFonts
        // Open a Document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

        // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
        // Each time a document is rendered.
        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_EmbedFullFonts(true);

        // The output PDF will be embedded with all fonts found in the document.
        System::String outputPath = outputDataDir + u"EmbeddedFontsInPDF.EmbedAllFonts_out.pdf";
        doc->Save(outputPath, options);
        // ExEnd:EmbeddAllFonts
    }

    void EmbedSubsetFonts(System::String const& inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:EmbeddSubsetFonts
        // Open a Document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

        // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false.
        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_EmbedFullFonts(false);

        // The output PDF will contain subsets of the fonts in the document. Only the glyphs used
        // In the document are included in the PDF fonts.
        System::String outputPath = outputDataDir + u"EmbeddedFontsInPDF.EmbeddSubsetFonts_out.pdf";
        doc->Save(outputPath, options);
        // ExEnd:EmbeddSubsetFonts
        std::cout << "Subset Fonts embedded successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetFontEmbeddingMode(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:SetFontEmbeddingMode
        // Load the document to render.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

        // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false.
        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_FontEmbeddingMode(PdfFontEmbeddingMode::EmbedNone);

        // The output PDF will be saved without embedding standard windows fonts.
        System::String outputPath = outputDataDir + u"EmbeddedFontsInPDF.DisableEmbedWindowsFonts_out.pdf";
        doc->Save(outputPath, options);
        // ExEnd:SetFontEmbeddingMode
    }
}

void EmbeddedFontsInPDF()
{
    std::cout << "EmbeddedFontsInPDF example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();
    
    EmbedAllFonts(inputDataDir, outputDataDir);
    EmbedSubsetFonts(inputDataDir, outputDataDir);
    SetFontEmbeddingMode(inputDataDir, outputDataDir);

    std::cout << "All Fonts embedded successfully." << std::endl << "File saved at " << outputDataDir.ToUtf8String() << std::endl;
    std::cout << "EmbeddedFontsInPDF example finished." << std::endl << std::endl;
}