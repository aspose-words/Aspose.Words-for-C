#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/WarningType.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Fonts/DefaultFontSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSubstitutionSettings.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;

namespace
{
    class HandleDocumentWarnings : public IWarningCallback
    {
        typedef HandleDocumentWarnings ThisType;
        typedef IWarningCallback BaseType;

        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);

    public:
        void Warning(System::SharedPtr<WarningInfo> info) override;
    };


    void HandleDocumentWarnings::Warning(System::SharedPtr<WarningInfo> info)
    {
        // We are only interested in fonts being substituted.
        if (info->get_WarningType() == WarningType::FontSubstitution)
        {
            std::cout << "Font substitution: " << info->get_Description().ToUtf8String() << std::endl;
        }
    }

    void ReceiveWarningNotification(System::SharedPtr<Document> doc, System::String const &outputDataDir)
    {
        // ExStart:ReceiveWarningNotification 
        // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occured during rendering
        // Are stored until the document save and then sent to the appropriate WarningCallback.
        doc->UpdatePageLayout();

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
        System::SharedPtr<HandleDocumentWarnings> callback = System::MakeObject<HandleDocumentWarnings>();

        doc->set_WarningCallback(callback);
        System::String outputPath = outputDataDir + u"ReceiveNotificationsOfFont.ReceiveWarningNotification.pdf";
        // Even though the document was rendered previously, any save warnings are notified to the user during document save.
        doc->Save(outputPath);
        // ExEnd:ReceiveWarningNotification  
    }
}

void ReceiveNotificationsOfFont()
{
    std::cout << "ReceiveNotificationsOfFont example started." << std::endl;
    // ExStart:ReceiveNotificationsOfFonts 
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

    System::SharedPtr<FontSettings> fontSettings = System::MakeObject<FontSettings>();

    // We can choose the default font to use in the case of any missing fonts.
    fontSettings->get_SubstitutionSettings()->get_DefaultFontSubstitution()->set_DefaultFontName(u"Arial");
    // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
    // Find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default
    // Font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback.
    fontSettings->SetFontsFolder(System::String::Empty, false);

    // Create a new class implementing IWarningCallback which collect any warnings produced during document save.
    System::SharedPtr<HandleDocumentWarnings> callback = System::MakeObject<HandleDocumentWarnings>();

    doc->set_WarningCallback(callback);
    // Set font settings
    doc->set_FontSettings(fontSettings);
    System::String outputPath = outputDataDir + u"ReceiveNotificationsOfFont.pdf";
    // Pass the save options along with the save path to the save method.
    doc->Save(outputPath);
    // ExEnd:ReceiveNotificationsOfFonts 
    std::cout << "Receive notifications of font substitutions by using IWarningCallback processed." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    ReceiveWarningNotification(doc, outputDataDir);
    std::cout << "ReceiveNotificationsOfFont example finished." << std::endl << std::endl;
}