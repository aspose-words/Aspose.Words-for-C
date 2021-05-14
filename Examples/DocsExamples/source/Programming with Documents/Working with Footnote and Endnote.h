#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Notes/EndnoteOptions.h>
#include <Aspose.Words.Cpp/Notes/EndnotePosition.h>
#include <Aspose.Words.Cpp/Notes/Footnote.h>
#include <Aspose.Words.Cpp/Notes/FootnoteNumberingRule.h>
#include <Aspose.Words.Cpp/Notes/FootnoteOptions.h>
#include <Aspose.Words.Cpp/Notes/FootnotePosition.h>
#include <Aspose.Words.Cpp/Notes/FootnoteType.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Notes;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithFootnotes : public DocsExamplesBase
{
public:
    void SetFootNoteColumns()
    {
        //ExStart:SetFootNoteColumns
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Specify the number of columns with which the footnotes area is formatted.
        doc->get_FootnoteOptions()->set_Columns(3);

        doc->Save(ArtifactsDir + u"WorkingWithFootnotes.SetFootNoteColumns.docx");
        //ExEnd:SetFootNoteColumns
    }

    void SetFootnoteAndEndNotePosition()
    {
        //ExStart:SetFootnoteAndEndNotePosition
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        doc->get_FootnoteOptions()->set_Position(FootnotePosition::BeneathText);
        doc->get_EndnoteOptions()->set_Position(EndnotePosition::EndOfSection);

        doc->Save(ArtifactsDir + u"WorkingWithFootnotes.SetFootnoteAndEndNotePosition.docx");
        //ExEnd:SetFootnoteAndEndNotePosition
    }

    void SetEndnoteOptions()
    {
        //ExStart:SetEndnoteOptions
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Some text");
        builder->InsertFootnote(FootnoteType::Endnote, u"Footnote text.");

        SharedPtr<EndnoteOptions> option = doc->get_EndnoteOptions();
        option->set_RestartRule(FootnoteNumberingRule::RestartPage);
        option->set_Position(EndnotePosition::EndOfSection);

        doc->Save(ArtifactsDir + u"WorkingWithFootnotes.SetEndnoteOptions.docx");
        //ExEnd:SetEndnoteOptions
    }
};

}} // namespace DocsExamples::Programming_with_Documents
