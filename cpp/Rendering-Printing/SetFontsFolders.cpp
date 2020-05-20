#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/array.h>
#include <Aspose.Words.Cpp/Model/Fonts/SystemFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;

void SetFontsFolders()
{
    std::cout << "SetFontsFolders example started." << std::endl;

    typedef System::SharedPtr<FontSourceBase> FontSourceBasePtr;

    // ExStart:SetFontsFolders
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    FontSettings::get_DefaultInstance()->SetFontsSources(System::MakeArray<FontSourceBasePtr>({System::MakeObject<SystemFontSource>(), System::MakeObject<FolderFontSource>(u"C:\\MyFonts\\", true)}));

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");
    const System::String outputPath = outputDataDir + u"Rendering.SetFontsFolders_out.pdf";
    doc->Save(outputPath);
    // ExEnd:SetFontsFolders

    std::cout << "Document saved as PDF successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "SetFontsFolders example finished." << std::endl << std::endl;
}
