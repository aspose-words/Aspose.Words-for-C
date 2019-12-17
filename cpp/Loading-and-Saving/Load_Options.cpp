#include "stdafx.h"
#include "examples.h"

#include <Model/Bookmarks/BookmarkStart.h>
#include <Model/Bookmarks/BookmarkEnd.h>
#include <Model/Document/FileFormatInfo.h>
#include <Model/Document/FileFormatUtil.h>
#include <Model/Document/LoadOptions.h>
#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/SaveFormat.h>
#include <Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Saving/OdtSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Saving;

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

    void AnnotationsAtBlockLevel(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:AnnotationsAtBlockLevel
        System::SharedPtr<LoadOptions> options = System::MakeObject<LoadOptions>();
        options->set_AnnotationsAtBlockLevel(true);
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"AnnotationsAtBlockLevel.docx", options);
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::SharedPtr<StructuredDocumentTag> sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChildNodes(NodeType::StructuredDocumentTag, true)->idx_get(0));

        System::SharedPtr<BookmarkStart> start = builder->StartBookmark(u"bm");
        System::SharedPtr<BookmarkEnd> end = builder->EndBookmark(u"bm");

        sdt->get_ParentNode()->InsertBefore(start, sdt);
        sdt->get_ParentNode()->InsertAfter(end, sdt);

        System::String outputPath = outputDataDir + u"Load_Options.AnnotationsAtBlockLevel.docx";
        //Save the document into DOCX
        doc->Save(outputPath, SaveFormat::Docx);
        // ExEnd:AnnotationsAtBlockLevel
        std::cout << "AnnotationsAtBlockLevel executed successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void Load_Options()
{
    std::cout << "Load_Options example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_QuickStart();
    System::String outputDataDir = GetOutputDataDir_QuickStart();
    LoadOptionsUpdateDirtyFields(inputDataDir, outputDataDir);
    // TODO (std_string): doesn't work properly, but raises FileCorruptedException instead of NotSupportedException / NotImplemented Exception
    //LoadAndSaveEncryptedODT(inputDataDir, outputDataDir);
    VerifyODTdocument(inputDataDir);
    ConvertShapeToOfficeMath(inputDataDir, outputDataDir);
    AnnotationsAtBlockLevel(inputDataDir, outputDataDir);
    std::cout << "Load_Options example finished." << std::endl << std::endl;
}