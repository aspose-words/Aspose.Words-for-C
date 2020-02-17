#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;

void SetTrueTypeFontsFolder()
{
    std::cout << "SetTrueTypeFontsFolder example started." << std::endl;
    // ExStart:SetTrueTypeFontsFolder
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

    System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>();

    // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
    // Fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
    // FontSettings.SetFontSources instead.
    fontSettings->SetFontsFolder(u"C:\\MyFonts\\", false);
    // Set font settings
    doc->set_FontSettings(fontSettings);
    System::String outputPath = outputDataDir + u"SetTrueTypeFontsFolder.pdf";
    doc->Save(outputPath);
    // ExEnd:SetTrueTypeFontsFolder
    std::cout << "True type fonts folder setup successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "SetTrueTypeFontsFolder example finished." << std::endl << std::endl;
}