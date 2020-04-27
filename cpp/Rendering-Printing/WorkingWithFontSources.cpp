#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Model/Fonts/PhysicalFontInfo.h>
#include <Aspose.Words.Cpp/Model/Fonts/SystemFontSource.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;

namespace
{
    typedef System::SharedPtr<FontSourceBase> TFontSourceBasePtr;
    typedef System::Collections::Generic::List<TFontSourceBasePtr> TFontSourceBasePtrList;

    void GetListOfAvailableFonts(System::String const &inputDataDir)
    {
        // ExStart:GetListOfAvailableFonts
        // Get available system fonts
        for (auto fontInfo : System::IterateOver(System::MakeObject<SystemFontSource>()->GetAvailableFonts()))
        {
            std::cout << "FontFamilyName : " <<  fontInfo->get_FontFamilyName().ToUtf8String() << std::endl;
            std::cout << "FullFontName  : " <<  fontInfo->get_FullFontName().ToUtf8String() << std::endl;
            std::cout << "Version  : " <<  fontInfo->get_Version().ToUtf8String() << std::endl;
            std::cout << "FilePath : " <<  fontInfo->get_FilePath().ToUtf8String() << std::endl;
        }

        // Get available fonts in folder
        for (auto fontInfo : System::IterateOver(System::MakeObject<FolderFontSource>(inputDataDir, true)->GetAvailableFonts()))
        {
            std::cout << "FontFamilyName : " <<  fontInfo->get_FontFamilyName().ToUtf8String() << std::endl;
            std::cout << "FullFontName  : " <<  fontInfo->get_FullFontName().ToUtf8String() << std::endl;
            std::cout << "Version  : " <<  fontInfo->get_Version().ToUtf8String() << std::endl;
            std::cout << "FilePath : " <<  fontInfo->get_FilePath().ToUtf8String() << std::endl;
        }

        // Get available fonts from FontSettings
        for (System::SharedPtr<FontSourceBase> fontsSource : FontSettings::get_DefaultInstance()->GetFontsSources())
        {
            for (auto fontInfo : System::IterateOver(fontsSource->GetAvailableFonts()))
            {
                std::cout << "FontFamilyName : " <<  fontInfo->get_FontFamilyName().ToUtf8String() << std::endl;
                std::cout << "FullFontName  : " <<  fontInfo->get_FullFontName().ToUtf8String() << std::endl;
                std::cout << "Version  : " <<  fontInfo->get_Version().ToUtf8String() << std::endl;
                std::cout << "FilePath : " <<  fontInfo->get_FilePath().ToUtf8String() << std::endl;
            }
        }
        // ExEnd:GetListOfAvailableFonts   
    }

    void SetFontsFolder(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:SetFontsFolder
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");
        System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>();

        // Retrieve the array of environment-dependent font sources that are searched by default. For example this will contain a "Windows\Fonts\" source on a Windows machines.
        // We add this array to a new ArrayList to make adding or removing font entries much easier.
        System::SharedPtr<TFontSourceBasePtrList> fontSources = System::MakeObject<TFontSourceBasePtrList>(fontSettings->GetFontsSources());

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        System::SharedPtr<FolderFontSource> folderFontSource = System::MakeObject<FolderFontSource>(u"C:\\MyFonts\\", true);
        
        // Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources->Add(folderFontSource);

        // Convert the Arraylist of source back into a primitive array of FontSource objects.
        System::ArrayPtr<TFontSourceBasePtr> updatedFontSources = fontSources->ToArray();

        // Apply the new set of font sources to use.
        fontSettings->SetFontsSources(updatedFontSources);
        // Set font settings
        doc->set_FontSettings(fontSettings);

        System::String outputPath = outputDataDir + u"SetFontsFoldersSystemAndCustomFolder.pdf";
        doc->Save(outputPath);
        // ExEnd:SetFontsFolder
    }
}

void WorkingWithFontSources()
{
    std::cout << "WorkingWithFontSources example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    GetListOfAvailableFonts(inputDataDir);
    SetFontsFolder(inputDataDir, outputDataDir);
    std::cout << "WorkingWithFontSources example finished." << std::endl << std::endl;
}