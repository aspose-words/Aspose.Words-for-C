#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
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

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace
{
    //ExStart:DocumentLoadingWarningCallback
    class DocumentLoadingWarningCallback : public IWarningCallback
    {
        public:
            System::SharedPtr<WarningInfoCollection> mWarnings;
            void Warning(System::SharedPtr<WarningInfo> info);
            DocumentLoadingWarningCallback();

        protected:
            System::Object::shared_members_type GetSharedMembers() override;
    };

    //RTTI_INFO_IMPL_HASH(1974866485u, DocumentLoadingWarningCallback, ThisTypeBaseTypesInfo);

    void DocumentLoadingWarningCallback::Warning(System::SharedPtr<WarningInfo> info)
    {
        // Prints warnings and their details as they arise during document loading.
        std::cout << "WARNING: {info.WarningType}, source: {info.Source}" << std::endl;
        std::cout << "\tDescription: {info.Description}" << std::endl;
    }

    DocumentLoadingWarningCallback::DocumentLoadingWarningCallback() : mWarnings(System::MakeObject<WarningInfoCollection>())
    {
    }

    System::Object::shared_members_type DocumentLoadingWarningCallback::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("DocumentLoadingWarningCallback::mWarnings", this->mWarnings);
        return result;
    }
    //ExEnd:DocumentLoadingWarningCallback
    
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
        loadOptions->set_MswVersion(MsWordVersion::Word2003);
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"document.doc", loadOptions);

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
    LoadOptionsEncoding(inputDataDir);
    LoadOptionsWarningCallback(inputDataDir);
    SetTempFolder(inputDataDir);
    std::cout << "Load_Options example finished." << std::endl << std::endl;
}