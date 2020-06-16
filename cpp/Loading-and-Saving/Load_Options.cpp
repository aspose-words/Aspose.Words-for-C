#include "stdafx.h"
#include "examples.h"
#include <system/type_info.h>
#include <system/convert.h>
#include <system/enumerator_adapter.h>
#include <drawing/image_converter.h>
#include <system/component_model/type_converter.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/WarningType.h>
#include <Aspose.Words.Cpp/Model/Loading/IResourceLoadingCallback.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingAction.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingArgs.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
using namespace Aspose::Words::Loading;
using namespace Aspose::Warnings;

namespace
{
    //ExStart:DocumentLoadingWarningCallback
    class DocumentLoadingWarningCallback : public IWarningCallback
    {
        public:
            System::SharedPtr<WarningInfoCollection> mWarnings;
            void Warning(System::SharedPtr<WarningInfo> info) override;
            DocumentLoadingWarningCallback();
    };

    void DocumentLoadingWarningCallback::Warning(System::SharedPtr<WarningInfo> info)
    {
        // Prints warnings and their details as they arise during document loading.
        std::cout << "WARNING: {info->get_WarningType} " << std::endl;
        std::cout << "Source: {info->get_Source} " << std::endl;
        std::cout << "\tDescription: {info->get_Description} " << std::endl;
    }

    DocumentLoadingWarningCallback::DocumentLoadingWarningCallback() : mWarnings(System::MakeObject<WarningInfoCollection>())
    {
    }
    //ExEnd:DocumentLoadingWarningCallback

    //ExStart:HtmlLinkedResourceLoadingCallback
    class HtmlLinkedResourceLoadingCallback : public IResourceLoadingCallback
    {
        public:
            ResourceLoadingAction ResourceLoading(System::SharedPtr<ResourceLoadingArgs> args) override;
            HtmlLinkedResourceLoadingCallback();
    };

    HtmlLinkedResourceLoadingCallback::HtmlLinkedResourceLoadingCallback()
    {
    }

    ResourceLoadingAction HtmlLinkedResourceLoadingCallback::ResourceLoading(System::SharedPtr<ResourceLoadingArgs> args)
    {
        switch (args->get_ResourceType())
        {
            case ResourceType::CssStyleSheet:
            {
                std::cout << "External CSS Stylesheet found upon loading: " << args->get_OriginalUri().ToUtf8String() << std::endl;

                // CSS file will don't used in the document
                return ResourceLoadingAction::Skip;
            }
            case ResourceType::Image:
            {
                // Replaces all images with a substitute
                System::String newImageFilename = u"Logo.jpg";
                std::cout << "\tImage will be substituted with: " << newImageFilename.ToUtf8String() << std::endl;

                System::SharedPtr<System::Drawing::Image> newImage = System::Drawing::Image::FromFile(GetInputDataDir_LoadingAndSaving() + newImageFilename);
                System::SharedPtr<System::Drawing::ImageConverter> converter = System::MakeObject<System::Drawing::ImageConverter>();
                auto imageBytes = System::DynamicCast<System::Array<uint8_t>>(converter->ConvertTo(nullptr, nullptr, newImage, System::ObjectExt::GetType<System::Array<uint8_t>>()));
                //System::ArrayPtr<uint8_t> imageBytes = System::IO::File::ReadAllBytes(GetInputDataDir_LoadingAndSaving() + newImageFilename);
                args->SetData(imageBytes);

                // New images will be used instead of presented in the document
                return ResourceLoadingAction::UserProvided;
            }
            case ResourceType::Document:
            {
                std::cout << "External document found upon loading: " << args->get_OriginalUri().ToUtf8String() << std::endl;

                // Will be used as usual
                return ResourceLoadingAction::Default;
            }
            default:
                throw System::InvalidOperationException(u"Unexpected ResourceType value.");
        }
    }
    
    //ExEnd:HtmlLinkedResourceLoadingCallback


    void LoadOptionsWarningCallback(System::String const& inputDataDir)
    {
        //ExStart:LoadOptionsWarningCallback
        // Create a new LoadOptions object and set its WarningCallback property. 
        System::SharedPtr<LoadOptions> loadOptions = System::MakeObject<LoadOptions>();

        System::SharedPtr<DocumentLoadingWarningCallback> callback = System::MakeObject<DocumentLoadingWarningCallback>();
        loadOptions->set_WarningCallback(callback);

        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"document.docx", loadOptions);
        //ExEnd:LoadOptionsWarningCallback
    }

    void LoadOptionsUpdateDirtyFields(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:LoadOptionsUpdateDirtyFields
        System::SharedPtr<LoadOptions> lo = System::MakeObject<LoadOptions>();

        //Update the fields with the dirty attribute
        lo->set_UpdateDirtyFields(true);

        //Load the Word document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"input.docx", lo);

        System::String outputPath = outputDataDir + u"Load_Options.LoadOptionsUpdateDirtyFields.docx";
        //Save the document into DOCX
        doc->Save(outputPath, SaveFormat::Docx);
        // ExEnd:LoadOptionsUpdateDirtyFields
        std::cout << "Update the fields with the dirty attribute successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void LoadAndSaveEncryptedODT(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:LoadAndSaveEncryptedODT
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"encrypted.odt", System::MakeObject<LoadOptions>(u"password"));

        System::String outputPath = outputDataDir + u"Load_Options.LoadAndSaveEncryptedODT.odt";
        doc->Save(outputPath, System::MakeObject<OdtSaveOptions>(u"newpassword"));
        // ExEnd:LoadAndSaveEncryptedODT 
        std::cout << "Load and save encrypted document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void VerifyODTdocument(System::String const &inputDataDir)
    {
        // ExStart:VerifyODTdocument
        System::SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(inputDataDir + u"encrypted.odt");
        std::cout << info->get_IsEncrypted() << std::endl;
        // ExEnd:VerifyODTdocument
        std::cout << "Verified encrypted document successfully." << std::endl << std::endl;
    }

    void ConvertShapeToOfficeMath(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:ConvertShapeToOfficeMath
        System::SharedPtr<LoadOptions> lo = System::MakeObject<LoadOptions>();
        lo->set_ConvertShapeToOfficeMath(true);

        // Specify load option to use previous default behaviour i.e. convert math shapes to office math ojects on loading stage.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"OfficeMath.docx", lo);
        System::String outputPath = outputDataDir + u"Load_Options.ConvertShapeToOfficeMath.docx";
        //Save the document into DOCX
        doc->Save(outputPath, SaveFormat::Docx);
        // ExEnd:ConvertShapeToOfficeMath
        std::cout << "Converted Shape To Office Math successfully." << std::endl << std::endl;
    }

    void SetMSWordVersion(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetMSWordVersion  
        System::SharedPtr<LoadOptions> loadOptions = System::MakeObject<LoadOptions>();
        // Change the loading version to Microsoft Word 2010.
        loadOptions->set_MswVersion(MsWordVersion::Word2010);
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"document.docx", loadOptions);

        System::String outputPath = outputDataDir + u"Load_Options.SetMSWordVersion.docx";
        doc->Save(outputPath);
        // ExEnd:SetMSWordVersion 
        std::cout << "Loaded with MS Word Version successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetTempFolder(System::String const& inputDataDir)
    {
        //ExStart:SetTempFolder
        // For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-C
        System::SharedPtr<LoadOptions> loadOptions = System::MakeObject<LoadOptions>();
        loadOptions->set_TempFolder(u"C:/TempFolder/");

        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"document.docx", loadOptions);
        //ExEnd:SetTempFolder
        std::cout << "Set Temp Folder successfully." << std::endl << std::endl;
    }

    void LoadOptionsEncoding(System::String const& inputDataDir)
    {
        //ExStart:LoadOptionsEncoding
        // Set the Encoding attribute in a LoadOptions object to override the automatically chosen encoding with the one we know to be correct
        System::SharedPtr<LoadOptions> loadOptions = System::MakeObject<LoadOptions>();
        loadOptions->set_Encoding(System::Text::Encoding::get_UTF7());

        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Encoded in UTF-7.txt", loadOptions);
        //ExEnd:LoadOptionsEncoding
        std::cout << "Set Encoding successfully." << std::endl << std::endl;
    }

    void LoadOptionsResourceLoadingCallback(System::String const& inputDataDir)
    {
        //ExStart:LoadOptionsResourceLoadingCallback
        // Create a new LoadOptions object and set its ResourceLoadingCallback attribute as an instance of our IResourceLoadingCallback implementation
        System::SharedPtr<LoadOptions> loadOptions = System::MakeObject<LoadOptions>();
        loadOptions->set_ResourceLoadingCallback(System::MakeObject<HtmlLinkedResourceLoadingCallback>());

        // When we open an Html document, external resources such as references to CSS stylesheet files and external images
        // will be handled in a custom manner by the loading callback as the document is loaded
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Images.html", loadOptions);
        doc->Save(inputDataDir + u"Document.LoadOptionsCallback_out.pdf");
        //ExEnd:LoadOptionsResourceLoadingCallback
        std::cout << "Load Options Resource Loading Callback successfully." << std::endl << std::endl;
    }
}

void Load_Options()
{
    std::cout << "Load_Options example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();
    LoadOptionsUpdateDirtyFields(inputDataDir, outputDataDir);
    LoadAndSaveEncryptedODT(inputDataDir, outputDataDir);
    VerifyODTdocument(inputDataDir);
    ConvertShapeToOfficeMath(inputDataDir, outputDataDir);
    SetMSWordVersion(inputDataDir, outputDataDir);
    //LoadOptionsEncoding(inputDataDir);
    LoadOptionsWarningCallback(inputDataDir);
    SetTempFolder(inputDataDir);
    LoadOptionsResourceLoadingCallback(inputDataDir);
    std::cout << "Load_Options example finished." << std::endl << std::endl;
}