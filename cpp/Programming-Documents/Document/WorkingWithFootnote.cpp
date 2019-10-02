#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnotePosition.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteOptions.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteNumberingRule.h>
#include <Aspose.Words.Cpp/Model/Footnotes/EndnotePosition.h>
#include <Aspose.Words.Cpp/Model/Footnotes/EndnoteOptions.h>

using namespace Aspose::Words;

namespace
{
    void SetFootNoteColumns(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetFootNoteColumns
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.docx");
        //Specify the number of columns with which the footnotes area is formatted.
        doc->get_FootnoteOptions()->set_Columns(3);
        System::String outputPath = outputDataDir + u"WorkingWithFootnote.SetFootNoteColumns.docx";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetFootNoteColumns
        std::cout << "Footnote number of columns set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetFootnoteAndEndNotePosition(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetFootnoteAndEndNotePosition
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.docx");
        //Set footnote and endnode position.
        doc->get_FootnoteOptions()->set_Position(FootnotePosition::BeneathText);
        doc->get_EndnoteOptions()->set_Position(EndnotePosition::EndOfSection);
        System::String outputPath = outputDataDir + u"WorkingWithFootnote.SetFootnoteAndEndNotePosition.docx";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetFootnoteAndEndNotePosition
        std::cout << "Footnote number of columns set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetEndnoteOptions(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetEndnoteOptions
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.docx");
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Some text");
        builder->InsertFootnote(FootnoteType::Endnote, u"Eootnote text.");
        System::SharedPtr<EndnoteOptions> option = doc->get_EndnoteOptions();
        option->set_RestartRule(FootnoteNumberingRule::RestartPage);
        option->set_Position(EndnotePosition::EndOfSection);
        System::String outputPath = outputDataDir + u"WorkingWithFootnote.SetEndnoteOptions.docx";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetEndnoteOptions
        std::cout << "Endnote is inserted at the end of section successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkingWithFootnote()
{
    std::cout << "WorkingWithFootnote example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    SetFootNoteColumns(inputDataDir, outputDataDir);
    SetFootnoteAndEndNotePosition(inputDataDir, outputDataDir);
    SetEndnoteOptions(inputDataDir, outputDataDir);
    std::cout << "WorkingWithFootnote example finished." << std::endl << std::endl;
}