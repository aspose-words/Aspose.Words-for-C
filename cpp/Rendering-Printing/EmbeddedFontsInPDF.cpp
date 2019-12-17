#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Saving/PdfSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void EmbeddSubsetFonts(System::SharedPtr<Document> doc, System::String const &outputDataDir)
    {
        // ExStart:EmbeddSubsetFonts
        // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false.
        System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
        options->set_EmbedFullFonts(false);
        System::String outputPath = outputDataDir + u"EmbeddedFontsInPDF.EmbeddSubsetFonts.pdf";
        // The output PDF will contain subsets of the fonts in the document. Only the glyphs used
        // In the document are included in the PDF fonts.
        doc->Save(outputPath, options);
        // ExEnd:EmbeddSubsetFonts
        std::cout << "Subset Fonts embedded successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void EmbeddedFontsInPDF()
{
    std::cout << "EmbeddedFontsInPDF example started." << std::endl;
    // ExStart:EmbeddAllFonts
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

    // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
    // Each time a document is rendered.
    System::SharedPtr<PdfSaveOptions> options = System::MakeObject<PdfSaveOptions>();
    options->set_EmbedFullFonts(true);

    System::String outputPath = outputDataDir + u"EmbeddedFontsInPDF.pdf";
    // The output PDF will be embedded with all fonts found in the document.
    doc->Save(outputPath, options);
    // ExEnd:EmbeddAllFonts
    std::cout << "All Fonts embedded successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    EmbeddSubsetFonts(doc, outputDataDir);
    std::cout << "EmbeddedFontsInPDF example finished." << std::endl << std::endl;
}