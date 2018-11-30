#include "stdafx.h"
#include "examples.h"

#include "system/shared_ptr.h"

#include "Model/Document/Document.h"
#include <Model/Document/DocumentBuilder.h>
#include <Model/Footnotes/FootnotePosition.h>
#include <Model/Footnotes/FootnoteType.h>
#include <Model/Footnotes/FootnoteOptions.h>
#include <Model/Footnotes/FootnoteNumberingRule.h>
#include <Model/Footnotes/EndnotePosition.h>
#include <Model/Footnotes/EndnoteOptions.h>

using namespace Aspose::Words;

namespace
{
    void SetFootNoteColumns(System::String const &dataDir)
    {
        // ExStart:SetFootNoteColumns
        auto doc = System::MakeObject<Document>(dataDir + u"TestFile.docx");
        //Specify the number of columns with which the footnotes area is formatted.
        doc->get_FootnoteOptions()->set_Columns(3);
        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithFootnote.SetFootNoteColumns.docx");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetFootNoteColumns
        std::cout << "Footnote number of columns set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetFootnoteAndEndNotePosition(System::String const &dataDir)
    {
        // ExStart:SetFootnoteAndEndNotePosition
        auto doc = System::MakeObject<Document>(dataDir + u"TestFile.docx");
        //Set footnote and endnode position.
        doc->get_FootnoteOptions()->set_Position(FootnotePosition::BeneathText);
        doc->get_EndnoteOptions()->set_Position(EndnotePosition::EndOfSection);
        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithFootnote.SetFootnoteAndEndNotePosition.docx");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetFootnoteAndEndNotePosition
        std::cout << "Footnote number of columns set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetEndnoteOptions(System::String const &dataDir)
    {
        // ExStart:SetEndnoteOptions
        auto doc = System::MakeObject<Document>(dataDir + u"TestFile.docx");
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Some text");
        builder->InsertFootnote(FootnoteType::Endnote, u"Eootnote text.");
        auto option = doc->get_EndnoteOptions();
        option->set_RestartRule(FootnoteNumberingRule::RestartPage);
        option->set_Position(EndnotePosition::EndOfSection);
        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithFootnote.SetEndnoteOptions.docx");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetEndnoteOptions
        std::cout << "Endnote is inserted at the end of section successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkingWithFootnote()
{
    std::cout << "WorkingWithFootnote example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    SetFootNoteColumns(dataDir);
    SetFootnoteAndEndNotePosition(dataDir);
    SetEndnoteOptions(dataDir);
    std::cout << "WorkingWithFootnote example finished." << std::endl << std::endl;
}