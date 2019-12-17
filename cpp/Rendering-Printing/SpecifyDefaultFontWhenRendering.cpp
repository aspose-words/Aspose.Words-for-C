#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Fonts/DefaultFontSubstitutionRule.h>
#include <Model/Fonts/FontSettings.h>
#include <Model/Fonts/FontSubstitutionSettings.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;

void SpecifyDefaultFontWhenRendering()
{
    std::cout << "SpecifyDefaultFontWhenRendering example started." << std::endl;
    // ExStart:SpecifyDefaultFontWhenRendering
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

    System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>();

    // If the default font defined here cannot be found during rendering then the closest font on the machine is used instead.
    fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial Unicode MS");
    // Set font settings
    doc->set_FontSettings(fontSettings);
    System::String outputPath = outputDataDir + u"SpecifyDefaultFontWhenRendering.pdf";
    // Now the set default font is used in place of any missing fonts during any rendering calls.
    doc->Save(outputPath);
    // ExEnd:SpecifyDefaultFontWhenRendering 
    std::cout << "Default font is setup during rendering." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "SpecifyDefaultFontWhenRendering example finished." << std::endl << std::endl;
}