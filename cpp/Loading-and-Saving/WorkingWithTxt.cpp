#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/TxtLeadingSpacesOptions.h>
#include <Model/Document/TxtLoadOptions.h>
#include <Model/Document/TxtTrailingSpacesOptions.h>
#include <Model/Saving/TxtListIndentation.h>
#include <Model/Saving/TxtSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void SaveAsTxt(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:SaveAsTxt
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
        System::String outputPath = outputDataDir + u"WorkingWithTxt.SaveAsTxt.txt";
        doc->Save(outputPath);
        //ExEnd:SaveAsTxt
        std::cout << "Document saved as TXT." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void AddBidiMarks(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:AddBidiMarks
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Input.docx");
        System::SharedPtr<TxtSaveOptions> saveOptions = System::MakeObject<TxtSaveOptions>();
        saveOptions->set_AddBidiMarks(true);

        System::String outputPath = outputDataDir + u"WorkingWithTxt.AddBidiMarks.txt";
        doc->Save(outputPath, saveOptions);
        //ExEnd:AddBidiMarks
        std::cout << "Add bi-directional marks set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void DetectNumberingWithWhitespaces(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:DetectNumberingWithWhitespaces
        System::SharedPtr<TxtLoadOptions> loadOptions = System::MakeObject<TxtLoadOptions>();
        loadOptions->set_DetectNumberingWithWhitespaces(false);

        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"LoadTxt.txt", loadOptions);

        System::String outputPath = outputDataDir + u"WorkingWithTxt.DetectNumberingWithWhitespaces.docx";
        doc->Save(outputPath);
        //ExEnd:DetectNumberingWithWhitespaces
        std::cout << "Detect number with whitespaces successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void HandleSpacesOptions(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:HandleSpacesOptions
        System::SharedPtr<TxtLoadOptions> loadOptions = System::MakeObject<TxtLoadOptions>();

        loadOptions->set_LeadingSpacesOptions(TxtLeadingSpacesOptions::Trim);
        loadOptions->set_TrailingSpacesOptions(TxtTrailingSpacesOptions::Trim);
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"LoadTxt.txt", loadOptions);

        System::String outputPath = outputDataDir + u"WorkingWithTxt.HandleSpacesOptions.docx";
        doc->Save(outputPath);
        //ExEnd:HandleSpacesOptions
        std::cout << "Trim leading and trailing spaces while importing text document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ExportHeadersFootersMode(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:ExportHeadersFootersMode
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TxtExportHeadersFootersMode.docx");

        System::SharedPtr<TxtSaveOptions> options = System::MakeObject<TxtSaveOptions>();
        options->set_SaveFormat(SaveFormat::Text);

        // All headers and footers are placed at the very end of the output document.
        options->set_ExportHeadersFootersMode(TxtExportHeadersFootersMode::AllAtEnd);
        doc->Save(outputDataDir + u"WorkingWithTxt.ExportHeadersFootersMode.AllAtEnd.txt", options);

        // Only primary headers and footers are exported at the beginning and end of each section.
        options->set_ExportHeadersFootersMode(TxtExportHeadersFootersMode::PrimaryOnly);
        doc->Save(outputDataDir + u"WorkingWithTxt.ExportHeadersFootersMode.PrimaryOnly.txt", options);

        // No headers and footers are exported.
        options->set_ExportHeadersFootersMode(TxtExportHeadersFootersMode::None);
        doc->Save(outputDataDir + u"WorkingWithTxt.ExportHeadersFootersMode.None.txt", options);

        //ExEnd:ExportHeadersFootersMode
        std::cout << "Export text files with TxtExportHeadersFootersMode." << std::endl << "Files saved at " << outputDataDir.ToUtf8String() << std::endl;
    }

    void UseTabCharacterPerLevelForListIndentation(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:UseTabCharacterPerLevelForListIndentation
        // TODO (std_string) : think about document named input_document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Input.docx");

        System::SharedPtr<TxtSaveOptions> options = System::MakeObject<TxtSaveOptions>();
        options->get_ListIndentation()->set_Count(1);
        options->get_ListIndentation()->set_Character(u'\t');

        doc->Save(outputDataDir + u"WorkingWithTxt.UseTabCharacterPerLevelForListIndentation.txt", options);
        //ExEnd:UseTabCharacterPerLevelForListIndentation
    }

    void UseSpaceCharacterPerLevelForListIndentation(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:UseSpaceCharacterPerLevelForListIndentation
        // TODO (std_string) : think about document named input_document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Input.docx");

        System::SharedPtr<TxtSaveOptions> options = System::MakeObject<TxtSaveOptions>();
        options->get_ListIndentation()->set_Count(3);
        options->get_ListIndentation()->set_Character(u' ');

        doc->Save(outputDataDir + u"WorkingWithTxt.UseSpaceCharacterPerLevelForListIndentation.txt", options);
        //ExEnd:UseSpaceCharacterPerLevelForListIndentation
    }

    void DefaultLevelForListIndentation(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:DefaultLevelForListIndentation
        // TODO (std_string) : think about document named input_document
        System::SharedPtr<Document> doc1 = System::MakeObject<Document>(inputDataDir + u"Input.docx");
        doc1->Save(outputDataDir + u"WorkingWithTxt.DefaultLevelForListIndentation1.txt");

        //Document doc2 = new Document("input_document");
        System::SharedPtr<Document> doc2 = System::MakeObject<Document>(inputDataDir + u"Input.docx");
        System::SharedPtr<TxtSaveOptions> options = System::MakeObject<TxtSaveOptions>();
        doc2->Save(outputDataDir + u"WorkingWithTxt.DefaultLevelForListIndentation2.txt", options);
        //ExEnd:DefaultLevelForListIndentation
    }

}

void WorkingWithTxt()
{
    std::cout << "WorkingWithTxt example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

    SaveAsTxt(inputDataDir, outputDataDir);
    AddBidiMarks(inputDataDir, outputDataDir);
    DetectNumberingWithWhitespaces(inputDataDir, outputDataDir);
    HandleSpacesOptions(inputDataDir, outputDataDir);
    ExportHeadersFootersMode(inputDataDir, outputDataDir);
    UseTabCharacterPerLevelForListIndentation(inputDataDir, outputDataDir);
    UseSpaceCharacterPerLevelForListIndentation(inputDataDir, outputDataDir);
    DefaultLevelForListIndentation(inputDataDir, outputDataDir);
    std::cout << "WorkingWithTxt example finished." << std::endl << std::endl;
}