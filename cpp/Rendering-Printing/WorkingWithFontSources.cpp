#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Model/Fonts/PhysicalFontInfo.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;

namespace
{
    typedef System::SharedPtr<FontSourceBase> TFontSourceBasePtr;
    typedef System::Collections::Generic::List<TFontSourceBasePtr> TFontSourceBasePtrList;

    void GetListOfAvailableFonts(System::String const &inputDataDir)
    {
        // ExStart:GetListOfAvailableFonts
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.docx");

        System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>();
        System::SharedPtr<TFontSourceBasePtrList> fontSources = System::MakeObject<TFontSourceBasePtrList>(fontSettings->GetFontsSources());

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        System::SharedPtr<FolderFontSource> folderFontSource = System::MakeObject<FolderFontSource>(inputDataDir, true);

        // Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources->Add(folderFontSource);

        // Convert the Arraylist of source back into a primitive array of FontSource objects.
        System::ArrayPtr<TFontSourceBasePtr> updatedFontSources = fontSources->ToArray();

        for (System::SharedPtr<PhysicalFontInfo> fontInfo : System::IterateOver(updatedFontSources[0]->GetAvailableFonts()))
        {
            std::cout << "FontFamilyName : " << fontInfo->get_FontFamilyName().ToUtf8String() << std::endl;
            std::cout << "FullFontName  : " << fontInfo->get_FullFontName().ToUtf8String() << std::endl;
            std::cout << "Version  : " << fontInfo->get_Version().ToUtf8String() << std::endl;
            std::cout << "FilePath : " << fontInfo->get_FilePath().ToUtf8String() << std::endl;
        }
        // ExEnd:GetListOfAvailableFonts
    }
}

void WorkingWithFontSources()
{
    std::cout << "WorkingWithFontSources example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    GetListOfAvailableFonts(inputDataDir);
    std::cout << "WorkingWithFontSources example finished." << std::endl << std::endl;
}