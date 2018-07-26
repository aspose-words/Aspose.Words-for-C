#include "../../examples.h"

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
        auto doc = System::MakeObject<Document>(dataDir + u"OtherTestFile.doc");
        //Specify the number of columns with which the footnotes area is formatted.
        doc->get_FootnoteOptions()->set_Columns(3);
        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithFootnote.SetFootNoteColumns.doc");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetFootNoteColumns
        std::cout << "\nFootnote number of columns set successfully.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetFootnoteAndEndNotePosition(System::String const &dataDir)
    {
        // ExStart:SetFootnoteAndEndNotePosition
        auto doc = System::MakeObject<Document>(dataDir + u"OtherTestFile.doc");
        //Set footnote and endnode position.
        doc->get_FootnoteOptions()->set_Position(Aspose::Words::FootnotePosition::BeneathText);
        doc->get_EndnoteOptions()->set_Position(Aspose::Words::EndnotePosition::EndOfSection);
        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithFootnote.SetFootnoteAndEndNotePosition.doc");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetFootnoteAndEndNotePosition
        std::cout << "\nFootnote number of columns set successfully.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetEndnoteOptions(System::String const &dataDir)
    {
        // ExStart:SetEndnoteOptions
        auto doc = System::MakeObject<Document>(dataDir + u"OtherTestFile.doc");
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Some text");
        builder->InsertFootnote(Aspose::Words::FootnoteType::Endnote, u"Eootnote text.");
        auto option = doc->get_EndnoteOptions();
        option->set_RestartRule(Aspose::Words::FootnoteNumberingRule::RestartPage);
        option->set_Position(Aspose::Words::EndnotePosition::EndOfSection);
        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithFootnote.SetEndnoteOptions.doc");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetEndnoteOptions
        std::cout << "\nEootnote is inserted at the end of section successfully.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkingWithFootnote()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    SetFootNoteColumns(dataDir);
    SetFootnoteAndEndNotePosition(dataDir);
    SetEndnoteOptions(dataDir);
}