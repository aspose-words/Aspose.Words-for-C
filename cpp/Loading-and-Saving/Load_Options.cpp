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

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;

namespace
{
    void LoadOptionsUpdateDirtyFields(System::String const &dataDir)
    {
        // ExStart:LoadOptionsUpdateDirtyFields
        System::SharedPtr<LoadOptions> lo = System::MakeObject<LoadOptions>();

        //Update the fields with the dirty attribute
        lo->set_UpdateDirtyFields(true);

        //Load the Word document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"input.docx", lo);

        System::String outputPath = dataDir + GetOutputFilePath(u"Load_Options.LoadOptionsUpdateDirtyFields.docx");
        //Save the document into DOCX
        doc->Save(outputPath, SaveFormat::Docx);
        // ExEnd:LoadOptionsUpdateDirtyFields
        std::cout << "Update the fields with the dirty attribute successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void VerifyODTdocument(System::String const &dataDir)
    {
        // ExStart:VerifyODTdocument
        System::SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(dataDir + u"encrypted.odt");
        std::cout << info->get_IsEncrypted() << std::endl;
        // ExEnd:VerifyODTdocument
    }

    void ConvertShapeToOfficeMath(System::String const &dataDir)
    {
        // ExStart:ConvertShapeToOfficeMath
        System::SharedPtr<LoadOptions> lo = System::MakeObject<LoadOptions>();
        lo->set_ConvertShapeToOfficeMath(true);

        // Specify load option to use previous default behaviour i.e. convert math shapes to office math ojects on loading stage.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"OfficeMath.docx", lo);
        System::String outputPath = dataDir + GetOutputFilePath(u"Load_Options.ConvertShapeToOfficeMath.docx");
        //Save the document into DOCX
        doc->Save(outputPath, SaveFormat::Docx);
        // ExEnd:ConvertShapeToOfficeMath
    }

    void AnnotationsAtBlockLevel(System::String dataDir)
    {
        // ExStart:AnnotationsAtBlockLevel
        System::SharedPtr<LoadOptions> options = System::MakeObject<LoadOptions>();
        options->set_AnnotationsAtBlockLevel(true);
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"AnnotationsAtBlockLevel.docx", options);
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::SharedPtr<StructuredDocumentTag> sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChildNodes(NodeType::StructuredDocumentTag, true)->idx_get(0));

        System::SharedPtr<BookmarkStart> start = builder->StartBookmark(u"bm");
        System::SharedPtr<BookmarkEnd> end = builder->EndBookmark(u"bm");

        sdt->get_ParentNode()->InsertBefore(start, sdt);
        sdt->get_ParentNode()->InsertAfter(end, sdt);

        System::String outputPath = dataDir + GetOutputFilePath(u"Load_Options.AnnotationsAtBlockLevel.docx");
        //Save the document into DOCX
        doc->Save(outputPath, SaveFormat::Docx);
        // ExEnd:AnnotationsAtBlockLevel
        std::cout << "AnnotationsAtBlockLevel executed successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void Load_Options()
{
    std::cout << "Load_Options example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_QuickStart();
    LoadOptionsUpdateDirtyFields(dataDir);
    VerifyODTdocument(dataDir);
    ConvertShapeToOfficeMath(dataDir);
    AnnotationsAtBlockLevel(dataDir);
    std::cout << "Load_Options example finished." << std::endl << std::endl;
}