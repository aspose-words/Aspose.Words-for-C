#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FolderFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <system/collections/list.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;

namespace
{
    typedef System::SharedPtr<FontSourceBase> TFontSourceBasePtr;
    typedef System::Collections::Generic::List<TFontSourceBasePtr> TFontSourceBasePtrList;
}

namespace
{
    void SetFontsFolders(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        std::cout << "SetFontsFoldersSystemAndCustomFolder example started." << std::endl;
        // ExStart:SetFontsFoldersSystemAndCustomFolder
        // The path to the documents directories.

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
        // ExEnd:SetFontsFoldersSystemAndCustomFolder 
        std::cout << "Fonts system and coustom folder is setup." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
        std::cout << "SetFontsFoldersSystemAndCustomFolder example finished." << std::endl << std::endl;
    }

    //ExStart:GetSubstitutionWithoutSuffixes
    class DocumentSubstitutionWarnings : public IWarningCallback
    {
        using ThisType = DocumentSubstitutionWarnings;
        using BaseType = IWarningCallback;
        using ThisTypeBaseTypesInfo = ::System::BaseTypesInfo<BaseType>;
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);
    public:

        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        void Warning(System::SharedPtr<Aspose::Words::WarningInfo> info) override
        {
            // We are only interested in fonts being substituted.
            if (info->get_WarningType() == WarningType::FontSubstitution)
            {
                FontWarnings->Warning(info);
            }
        }
        System::SharedPtr<WarningInfoCollection> FontWarnings = System::MakeObject<WarningInfoCollection>();
    };

    void GetSubstitutionWithoutSuffixes(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        auto doc = System::MakeObject<Document>(inputDataDir + u"Get substitution without suffixes.docx");

        auto substitutionWarningHandler = System::MakeObject<DocumentSubstitutionWarnings>();
        doc->set_WarningCallback(substitutionWarningHandler);

        auto fontSources = FontSettings::get_DefaultInstance()->GetFontsSources();
        auto folderFontSource = System::MakeObject<FolderFontSource>(inputDataDir + u"Fonts/", true);
        fontSources->data().push_back(folderFontSource);

        FontSettings::get_DefaultInstance()->SetFontsSources(fontSources);

        doc->Save(outputDataDir + u"GetSubstitutionWithoutSuffixes_out.pdf");
    }
    //ExEnd:GetSubstitutionWithoutSuffixes
}

void SetFontsFoldersSystemAndCustomFolder()
{
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    SetFontsFolders(inputDataDir, outputDataDir);
    GetSubstitutionWithoutSuffixes(inputDataDir, outputDataDir);
}