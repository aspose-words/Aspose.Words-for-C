#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/TxtLeadingSpacesOptions.h>
#include <Aspose.Words.Cpp/Model/Document/TxtLoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/TxtTrailingSpacesOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/TxtListIndentation.h>
#include <Aspose.Words.Cpp/Model/Saving/TxtSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>

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


    void DocumentTextDirection(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:DocumentTextDirection
        auto loadOptions = System::MakeObject<TxtLoadOptions>();
        loadOptions->set_DocumentDirection(DocumentDirection::Auto);

        auto doc = System::MakeObject<Document>(inputDataDir + u"Hebrew text.txt", loadOptions);

        auto paragraph = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
        std::cout << paragraph->get_ParagraphFormat()->get_Bidi() << '\n';

        doc->Save(outputDataDir + u"WorkingWithTxtLoadOptions.DocumentTextDirection.docx");
        //ExEnd:DocumentTextDirection
    }

    void UseTabCharacterPerLevelForListIndentation(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:UseTabCharacterPerLevelForListIndentation
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);

        // Create a list with three levels of indentation.
        builder->get_ListFormat()->ApplyNumberDefault();
        builder->Writeln(u"Item 1");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->ListIndent();
        builder->Write(u"Item 3");

        System::SharedPtr<TxtSaveOptions> options = System::MakeObject<TxtSaveOptions>();
        options->get_ListIndentation()->set_Count(1);
        options->get_ListIndentation()->set_Character(u'\t');

        doc->Save(outputDataDir + u"WorkingWithTxt.UseTabCharacterPerLevelForListIndentation.txt", options);
        //ExEnd:UseTabCharacterPerLevelForListIndentation
    }

    void UseSpaceCharacterPerLevelForListIndentation(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:UseSpaceCharacterPerLevelForListIndentation
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);

        // Create a list with three levels of indentation.
        builder->get_ListFormat()->ApplyNumberDefault();
        builder->Writeln(u"Item 1");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->ListIndent();
        builder->Write(u"Item 3");

        System::SharedPtr<TxtSaveOptions> options = System::MakeObject<TxtSaveOptions>();
        options->get_ListIndentation()->set_Count(3);
        options->get_ListIndentation()->set_Character(u' ');

        doc->Save(outputDataDir + u"WorkingWithTxt.UseSpaceCharacterPerLevelForListIndentation.txt", options);
        //ExEnd:UseSpaceCharacterPerLevelForListIndentation
    }

    void DefaultLevelForListIndentation(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:DefaultLevelForListIndentation
        System::SharedPtr<Document> doc1 = System::MakeObject<Document>(inputDataDir + u"input_document");
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
    DocumentTextDirection(inputDataDir, outputDataDir);

    UseSpaceCharacterPerLevelForListIndentation(inputDataDir, outputDataDir);
    UseTabCharacterPerLevelForListIndentation(inputDataDir, outputDataDir);
#if 0
    // Source document is missing
    DefaultLevelForListIndentation(inputDataDir, outputDataDir);
#endif
    std::cout << "WorkingWithTxt example finished." << std::endl << std::endl;
}