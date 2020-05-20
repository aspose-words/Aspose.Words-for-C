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

void SetFontsFoldersWithPriority()
{
    std::cout << "SetFontsFoldersWithPriority example started." << std::endl;

    typedef System::SharedPtr<FontSourceBase> FontSourceBasePtr;

    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    // ExStart:SetFontsFoldersWithPriority
    FontSettings::get_DefaultInstance()->SetFontsSources(System::MakeArray<FontSourceBasePtr>({System::MakeObject<SystemFontSource>(), System::MakeObject<FolderFontSource>(u"C:\\MyFonts\\", true, 1)}));
    // ExEnd:SetFontsFoldersWithPriority

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");
    const System::String outputPath = outputDataDir + u"Rendering.SetFontsFoldersWithPriority_out.pdf";
    doc->Save(outputPath);

    std::cout << "Document saved as PDF successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "SetFontsFoldersWithPriority example finished." << std::endl << std::endl;
}
