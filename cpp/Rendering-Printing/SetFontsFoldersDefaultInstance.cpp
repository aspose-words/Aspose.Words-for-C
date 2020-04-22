#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;

void SetFontsFoldersDefaultInstance()
{
    std::cout << "SetFontsFoldersDefaultInstance example started." << std::endl;

    typedef System::SharedPtr<FontSourceBase> FontSourceBasePtr;

    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    // ExStart:SetFontsFoldersDefaultInstance
    FontSettings::get_DefaultInstance()->SetFontsFolder(u"C:\\MyFonts\\", true);
    // ExEnd:SetFontsFoldersDefaultInstance

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");
    const System::String outputPath = outputDataDir + u"Rendering.SetFontsFoldersDefaultInstance_out.pdf";
    doc->Save(outputPath);

    std::cout << "Document saved as PDF successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "SetFontsFoldersDefaultInstance example finished." << std::endl << std::endl;
}
