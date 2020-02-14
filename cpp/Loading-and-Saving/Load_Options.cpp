#include "stdafx.h"
#include "examples.h"

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

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace
{
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
}

void Load_Options()
{
    std::cout << "Load_Options example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_QuickStart();
    System::String outputDataDir = GetOutputDataDir_QuickStart();
    LoadOptionsUpdateDirtyFields(inputDataDir, outputDataDir);
    LoadAndSaveEncryptedODT(inputDataDir, outputDataDir);
    VerifyODTdocument(inputDataDir);
    ConvertShapeToOfficeMath(inputDataDir, outputDataDir);
    SetMSWordVersion(inputDataDir, outputDataDir);
    std::cout << "Load_Options example finished." << std::endl << std::endl;
}