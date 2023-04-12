#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeImporter.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Notes/Footnote.h>
#include <Aspose.Words.Cpp/Notes/FootnoteType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Replacing/FindReplaceDirection.h>
#include <Aspose.Words.Cpp/Replacing/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Replacing/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Replacing/ReplaceAction.h>
#include <Aspose.Words.Cpp/Replacing/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <drawing/color.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/collections/list.h>
#include <system/convert.h>
#include <system/date_time.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/string_builder.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Replacing;

namespace ApiExamples {

class ExRange : public ApiExampleBase
{
public:
    void Replace()
    {
        //ExStart
        //ExFor:Range.Replace(String, String)
        //ExSummary:Shows how to perform a find-and-replace text operation on the contents of a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Greetings, _FullName_!");

        // Perform a find-and-replace operation on our document's contents and verify the number of replacements that took place.
        int replacementCount = doc->get_Range()->Replace(u"_FullName_", u"John Doe");

        ASSERT_EQ(1, replacementCount);
        ASSERT_EQ(u"Greetings, John Doe!", doc->GetText().Trim());
        //ExEnd
    }

    void ReplaceMatchCase(bool matchCase)
    {
        //ExStart
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExFor:FindReplaceOptions
        //ExFor:FindReplaceOptions.MatchCase
        //ExSummary:Shows how to toggle case sensitivity when performing a find-and-replace operation.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Ruby bought a ruby necklace.");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        auto options = MakeObject<FindReplaceOptions>();

        // Set the "MatchCase" flag to "true" to apply case sensitivity while finding strings to replace.
        // Set the "MatchCase" flag to "false" to ignore character case while searching for text to replace.
        options->set_MatchCase(matchCase);

        doc->get_Range()->Replace(u"Ruby", u"Jade", options);

        ASSERT_EQ(matchCase ? String(u"Jade bought a ruby necklace.") : String(u"Jade bought a Jade necklace."), doc->GetText().Trim());
        //ExEnd
    }

    void ReplaceFindWholeWordsOnly(bool findWholeWordsOnly)
    {
        //ExStart
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExFor:FindReplaceOptions
        //ExFor:FindReplaceOptions.FindWholeWordsOnly
        //ExSummary:Shows how to toggle standalone word-only find-and-replace operations.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Jackson will meet you in Jacksonville.");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        auto options = MakeObject<FindReplaceOptions>();

        // Set the "FindWholeWordsOnly" flag to "true" to replace the found text if it is not a part of another word.
        // Set the "FindWholeWordsOnly" flag to "false" to replace all text regardless of its surroundings.
        options->set_FindWholeWordsOnly(findWholeWordsOnly);

        doc->get_Range()->Replace(u"Jackson", u"Louis", options);

        ASSERT_EQ(findWholeWordsOnly ? String(u"Louis will meet you in Jacksonville.") : String(u"Louis will meet you in Louisville."), doc->GetText().Trim());
        //ExEnd
    }

    void IgnoreDeleted(bool ignoreTextInsideDeleteRevisions)
    {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreDeleted
        //ExSummary:Shows how to include or ignore text inside delete revisions during a find-and-replace operation.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");
        builder->Writeln(u"Hello again!");

        // Start tracking revisions and remove the second paragraph, which will create a delete revision.
        // That paragraph will persist in the document until we accept the delete revision.
        doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
        doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->Remove();
        doc->StopTrackRevisions();

        ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_IsDeleteRevision());

        // We can use a "FindReplaceOptions" object to modify the find and replace process.
        auto options = MakeObject<FindReplaceOptions>();

        // Set the "IgnoreDeleted" flag to "true" to get the find-and-replace
        // operation to ignore paragraphs that are delete revisions.
        // Set the "IgnoreDeleted" flag to "false" to get the find-and-replace
        // operation to also search for text inside delete revisions.
        options->set_IgnoreDeleted(ignoreTextInsideDeleteRevisions);

        doc->get_Range()->Replace(u"Hello", u"Greetings", options);

        ASSERT_EQ(ignoreTextInsideDeleteRevisions ? String(u"Greetings world!\rHello again!") : String(u"Greetings world!\rGreetings again!"),
                  doc->GetText().Trim());
        //ExEnd
    }

    void IgnoreInserted(bool ignoreTextInsideInsertRevisions)
    {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreInserted
        //ExSummary:Shows how to include or ignore text inside insert revisions during a find-and-replace operation.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");

        // Start tracking revisions and insert a paragraph. That paragraph will be an insert revision.
        doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
        builder->Writeln(u"Hello again!");
        doc->StopTrackRevisions();

        ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_IsInsertRevision());

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        auto options = MakeObject<FindReplaceOptions>();

        // Set the "IgnoreInserted" flag to "true" to get the find-and-replace
        // operation to ignore paragraphs that are insert revisions.
        // Set the "IgnoreInserted" flag to "false" to get the find-and-replace
        // operation to also search for text inside insert revisions.
        options->set_IgnoreInserted(ignoreTextInsideInsertRevisions);

        doc->get_Range()->Replace(u"Hello", u"Greetings", options);

        ASSERT_EQ(ignoreTextInsideInsertRevisions ? String(u"Greetings world!\rHello again!") : String(u"Greetings world!\rGreetings again!"),
                  doc->GetText().Trim());
        //ExEnd
    }

    void IgnoreFields(bool ignoreTextInsideFields)
    {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreFields
        //ExSummary:Shows how to ignore text inside fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");
        builder->InsertField(u"QUOTE", u"Hello again!");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        auto options = MakeObject<FindReplaceOptions>();

        // Set the "IgnoreFields" flag to "true" to get the find-and-replace
        // operation to ignore text inside fields.
        // Set the "IgnoreFields" flag to "false" to get the find-and-replace
        // operation to also search for text inside fields.
        options->set_IgnoreFields(ignoreTextInsideFields);

        doc->get_Range()->Replace(u"Hello", u"Greetings", options);

        ASSERT_EQ(ignoreTextInsideFields ? String(u"Greetings world!\r\u0013QUOTE\u0014Hello again!\u0015")
                                         : String(u"Greetings world!\r\u0013QUOTE\u0014Greetings again!\u0015"),
                  doc->GetText().Trim());
        //ExEnd
    }

    void IgnoreFootnote(bool isIgnoreFootnotes)
    {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreFootnotes
        //ExSummary:Shows how to ignore footnotes during a find-and-replace operation.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
        builder->InsertFootnote(FootnoteType::Footnote, u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.");

        builder->InsertParagraph();

        builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
        builder->InsertFootnote(FootnoteType::Endnote, u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.");

        // Set the "IgnoreFootnotes" flag to "true" to get the find-and-replace
        // operation to ignore text inside footnotes.
        // Set the "IgnoreFootnotes" flag to "false" to get the find-and-replace
        // operation to also search for text inside footnotes.
        auto options = MakeObject<FindReplaceOptions>();
        options->set_IgnoreFootnotes(isIgnoreFootnotes);
        doc->get_Range()->Replace(u"Lorem ipsum", u"Replaced Lorem ipsum", options);
        //ExEnd

        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        for (const auto& para : System::IterateOver<Paragraph>(paragraphs))
        {
            ASSERT_EQ(u"Replaced Lorem ipsum", para->get_Runs()->idx_get(0)->get_Text());
        }

        SharedPtr<System::Collections::Generic::List<SharedPtr<Footnote>>> footnotes =
            doc->GetChildNodes(NodeType::Footnote, true)->LINQ_Cast<SharedPtr<Footnote>>()->LINQ_ToList();
        ASSERT_EQ(isIgnoreFootnotes ? String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.")
                                    : String(u"Replaced Lorem ipsum dolor sit amet, consectetur adipiscing elit."),
                  footnotes->idx_get(0)->ToString(SaveFormat::Text).Trim());
        ASSERT_EQ(isIgnoreFootnotes ? String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.")
                                    : String(u"Replaced Lorem ipsum dolor sit amet, consectetur adipiscing elit."),
                  footnotes->idx_get(1)->ToString(SaveFormat::Text).Trim());
    }

    void UpdateFieldsInRange()
    {
        //ExStart
        //ExFor:Range.UpdateFields
        //ExSummary:Shows how to update all the fields in a range.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u" DOCPROPERTY Category");
        builder->InsertBreak(BreakType::SectionBreakEvenPage);
        builder->InsertField(u" DOCPROPERTY Category");

        // The above DOCPROPERTY fields will display the value of this built-in document property.
        doc->get_BuiltInDocumentProperties()->set_Category(u"MyCategory");

        // If we update the value of a document property, we will need to update all the DOCPROPERTY fields to display it.
        ASSERT_EQ(String::Empty, doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
        ASSERT_EQ(String::Empty, doc->get_Range()->get_Fields()->idx_get(1)->get_Result());

        // Update all the fields that are in the range of the first section.
        doc->get_FirstSection()->get_Range()->UpdateFields();

        ASSERT_EQ(u"MyCategory", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
        ASSERT_EQ(String::Empty, doc->get_Range()->get_Fields()->idx_get(1)->get_Result());
        //ExEnd
    }

    void ReplaceWithString()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"This one is sad.");
        builder->Writeln(u"That one is mad.");

        auto options = MakeObject<FindReplaceOptions>();
        options->set_MatchCase(false);
        options->set_FindWholeWordsOnly(true);

        doc->get_Range()->Replace(u"sad", u"bad", options);

        doc->Save(ArtifactsDir + u"Range.ReplaceWithString.docx");
    }

    void ReplaceWithRegex()
    {
        //ExStart
        //ExFor:Range.Replace(Regex, String)
        //ExSummary:Shows how to replace all occurrences of a regular expression pattern with other text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"I decided to get the curtains in gray, ideal for the grey-accented room.");

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"gr(a|e)y"), u"lavender");

        ASSERT_EQ(u"I decided to get the curtains in lavender, ideal for the lavender-accented room.", doc->GetText().Trim());
        //ExEnd
    }

    //ExStart
    //ExFor:FindReplaceOptions.ReplacingCallback
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:ReplacingArgs.Replacement
    //ExFor:IReplacingCallback
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExSummary:Shows how to replace all occurrences of a regular expression pattern with another string, while tracking all such replacements.
    void ReplaceWithCallback()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(String(u"Our new location in New York City is opening tomorrow. ") + u"Hope to see all our NYC-based customers at the opening!");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        auto options = MakeObject<FindReplaceOptions>();

        // Set a callback that tracks any replacements that the "Replace" method will make.
        auto logger = MakeObject<ExRange::TextFindAndReplacementLogger>();
        options->set_ReplacingCallback(logger);

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"New York City|NYC"), u"Washington", options);

        ASSERT_EQ(String(u"Our new location in (Old value:\"New York City\") Washington is opening tomorrow. ") +
                      u"Hope to see all our (Old value:\"NYC\") Washington-based customers at the opening!",
                  doc->GetText().Trim());

        ASSERT_EQ(String(u"\"New York City\" converted to \"Washington\" 20 characters into a Run node.\r\n") +
                      u"\"NYC\" converted to \"Washington\" 42 characters into a Run node.",
                  logger->GetLog().Trim());
    }

    /// <summary>
    /// Maintains a log of every text replacement done by a find-and-replace operation
    /// and notes the original matched text's value.
    /// </summary>
    class TextFindAndReplacementLogger : public IReplacingCallback
    {
    public:
        String GetLog()
        {
            return mLog->ToString();
        }

        TextFindAndReplacementLogger() : mLog(MakeObject<System::Text::StringBuilder>())
        {
        }

    private:
        SharedPtr<System::Text::StringBuilder> mLog;

        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            mLog->AppendLine(String::Format(u"\"{0}\" converted to \"{1}\" ", args->get_Match()->get_Value(), args->get_Replacement()) +
                             String::Format(u"{0} characters into a {1} node.", args->get_MatchOffset(), args->get_MatchNode()->get_NodeType()));

            args->set_Replacement(String::Format(u"(Old value:\"{0}\") {1}", args->get_Match()->get_Value(), args->get_Replacement()));
            return ReplaceAction::Replace;
        }
    };
    //ExEnd

    //ExStart
    //ExFor:FindReplaceOptions.ApplyFont
    //ExFor:FindReplaceOptions.ReplacingCallback
    //ExFor:ReplacingArgs.GroupIndex
    //ExFor:ReplacingArgs.GroupName
    //ExFor:ReplacingArgs.Match
    //ExFor:ReplacingArgs.MatchOffset
    //ExSummary:Shows how to apply a different font to new content via FindReplaceOptions.
    void ConvertNumbersToHexadecimal()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(String(u"Numbers that the find-and-replace operation will convert to hexadecimal and highlight:\n") + u"123, 456, 789 and 17379.");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        auto options = MakeObject<FindReplaceOptions>();

        // Set the "HighlightColor" property to a background color that we want to apply to the operation's resulting text.
        options->get_ApplyFont()->set_HighlightColor(System::Drawing::Color::get_LightGray());

        auto numberHexer = MakeObject<ExRange::NumberHexer>();
        options->set_ReplacingCallback(numberHexer);

        int replacementCount = doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"[0-9]+"), u"", options);

        std::cout << numberHexer->GetLog() << std::endl;

        ASSERT_EQ(4, replacementCount);
        ASSERT_EQ(String(u"Numbers that the find-and-replace operation will convert to hexadecimal and highlight:\r") + u"0x7B, 0x1C8, 0x315 and 0x43E3.",
                  doc->GetText().Trim());
        auto isLightGray = [](SharedPtr<Run> r)
        {
            return r->get_Font()->get_HighlightColor().ToArgb() == System::Drawing::Color::get_LightGray().ToArgb();
        };

        ASSERT_EQ(4, doc->GetChildNodes(NodeType::Run, true)->LINQ_OfType<SharedPtr<Run>>()->LINQ_Count(isLightGray));
    }

    /// <summary>
    /// Replaces numeric find-and-replacement matches with their hexadecimal equivalents.
    /// Maintains a log of every replacement.
    /// </summary>
    class NumberHexer : public IReplacingCallback
    {
    public:
        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            mCurrentReplacementNumber++;

            int number = System::Convert::ToInt32(args->get_Match()->get_Value());

            args->set_Replacement(String::Format(u"0x{0:X}", number));

            mLog->AppendLine(String::Format(u"Match #{0}", mCurrentReplacementNumber));
            mLog->AppendLine(String::Format(u"\tOriginal value:\t{0}", args->get_Match()->get_Value()));
            mLog->AppendLine(String::Format(u"\tReplacement:\t{0}", args->get_Replacement()));
            mLog->AppendLine(String::Format(u"\tOffset in parent {0} node:\t{1}", args->get_MatchNode()->get_NodeType(), args->get_MatchOffset()));

            mLog->AppendLine(String::IsNullOrEmpty(args->get_GroupName()) ? String::Format(u"\tGroup index:\t{0}", args->get_GroupIndex())
                                                                          : String::Format(u"\tGroup name:\t{0}", args->get_GroupName()));

            return ReplaceAction::Replace;
        }

        String GetLog()
        {
            return mLog->ToString();
        }

        NumberHexer() : mCurrentReplacementNumber(0), mLog(MakeObject<System::Text::StringBuilder>())
        {
        }

    private:
        int mCurrentReplacementNumber;
        SharedPtr<System::Text::StringBuilder> mLog;
    };
    //ExEnd

    void ApplyParagraphFormat()
    {
        //ExStart
        //ExFor:FindReplaceOptions.ApplyParagraphFormat
        //ExFor:Range.Replace(String, String)
        //ExSummary:Shows how to add formatting to paragraphs in which a find-and-replace operation has found matches.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Every paragraph that ends with a full stop like this one will be right aligned.");
        builder->Writeln(u"This one will not!");
        builder->Write(u"This one also will.");

        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        ASSERT_EQ(ParagraphAlignment::Left, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Alignment());
        ASSERT_EQ(ParagraphAlignment::Left, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Alignment());
        ASSERT_EQ(ParagraphAlignment::Left, paragraphs->idx_get(2)->get_ParagraphFormat()->get_Alignment());

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        auto options = MakeObject<FindReplaceOptions>();

        // Set the "Alignment" property to "ParagraphAlignment.Right" to right-align every paragraph
        // that contains a match that the find-and-replace operation finds.
        options->get_ApplyParagraphFormat()->set_Alignment(ParagraphAlignment::Right);

        // Replace every full stop that is right before a paragraph break with an exclamation point.
        int count = doc->get_Range()->Replace(u".&p", u"!&p", options);

        ASSERT_EQ(2, count);
        ASSERT_EQ(ParagraphAlignment::Right, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Alignment());
        ASSERT_EQ(ParagraphAlignment::Left, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Alignment());
        ASSERT_EQ(ParagraphAlignment::Right, paragraphs->idx_get(2)->get_ParagraphFormat()->get_Alignment());
        ASSERT_EQ(String(u"Every paragraph that ends with a full stop like this one will be right aligned!\r") + u"This one will not!\r" +
                      u"This one also will!",
                  doc->GetText().Trim());
        //ExEnd
    }

    void DeleteSelection()
    {
        //ExStart
        //ExFor:Node.Range
        //ExFor:Range.Delete
        //ExSummary:Shows how to delete all the nodes from a range.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add text to the first section in the document, and then add another section.
        builder->Write(u"Section 1. ");
        builder->InsertBreak(BreakType::SectionBreakContinuous);
        builder->Write(u"Section 2.");

        ASSERT_EQ(u"Section 1. \fSection 2.", doc->GetText().Trim());

        // Remove the first section entirely by removing all the nodes
        // within its range, including the section itself.
        doc->get_Sections()->idx_get(0)->get_Range()->Delete();

        ASSERT_EQ(1, doc->get_Sections()->get_Count());
        ASSERT_EQ(u"Section 2.", doc->GetText().Trim());
        //ExEnd
    }

    void RangesGetText()
    {
        //ExStart
        //ExFor:Range
        //ExFor:Range.Text
        //ExSummary:Shows how to get the text contents of all the nodes that a range covers.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Hello world!");

        ASSERT_EQ(u"Hello world!", doc->get_Range()->get_Text().Trim());
        //ExEnd
    }

    //ExStart
    //ExFor:FindReplaceOptions.UseLegacyOrder
    //ExSummary:Shows how to change the searching order of nodes when performing a find-and-replace text operation.
    void UseLegacyOrder(bool useLegacyOrder)
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert three runs which we can search for using a regex pattern.
        // Place one of those runs inside a text box.
        builder->Writeln(u"[tag 1]");
        SharedPtr<Shape> textBox = builder->InsertShape(ShapeType::TextBox, 100, 50);
        builder->Writeln(u"[tag 2]");
        builder->MoveTo(textBox->get_FirstParagraph());
        builder->Write(u"[tag 3]");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        auto options = MakeObject<FindReplaceOptions>();

        // Assign a custom callback to the "ReplacingCallback" property.
        auto callback = MakeObject<ExRange::TextReplacementTracker>();
        options->set_ReplacingCallback(callback);

        // If we set the "UseLegacyOrder" property to "true", the
        // find-and-replace operation will go through all the runs outside of a text box
        // before going through the ones inside a text box.
        // If we set the "UseLegacyOrder" property to "false", the
        // find-and-replace operation will go over all the runs in a range in sequential order.
        options->set_UseLegacyOrder(useLegacyOrder);

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"\\[tag \\d*\\]"), u"", options);

        SharedPtr<System::Collections::Generic::List<String>> matches = MakeObject<System::Collections::Generic::List<String>>();
        matches->Add(u"[tag 1]");
        if (useLegacyOrder)
        {
            matches->Add(u"[tag 3]");
            matches->Add(u"[tag 2]");
        }
        else
        {
            matches->Add(u"[tag 2]");
            matches->Add(u"[tag 3]");
        }
        ASPOSE_ASSERT_EQ(matches, callback->get_Matches());
    }

    /// <summary>
    /// Records the order of all matches that occur during a find-and-replace operation.
    /// </summary>
    class TextReplacementTracker : public IReplacingCallback
    {
    public:
        SharedPtr<System::Collections::Generic::List<String>> get_Matches()
        {
            return mMatches;
        }

        TextReplacementTracker() : mMatches(MakeObject<System::Collections::Generic::List<String>>())
        {
        }

    private:
        SharedPtr<System::Collections::Generic::List<String>> mMatches;

        ReplaceAction Replacing(SharedPtr<ReplacingArgs> e) override
        {
            get_Matches()->Add(e->get_Match()->get_Value());
            return ReplaceAction::Replace;
        }
    };
    //ExEnd

    void UseSubstitutions(bool useSubstitutions)
    {
        //ExStart
        //ExFor:FindReplaceOptions.UseSubstitutions
        //ExSummary:Shows how to replace the text with substitutions.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"John sold a car to Paul.");
        builder->Writeln(u"Jane sold a house to Joe.");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        auto options = MakeObject<FindReplaceOptions>();

        // Set the "UseSubstitutions" property to "true" to get
        // the find-and-replace operation to recognize substitution elements.
        // Set the "UseSubstitutions" property to "false" to ignore substitution elements.
        options->set_UseSubstitutions(useSubstitutions);

        auto regex = MakeObject<System::Text::RegularExpressions::Regex>(u"([A-z]+) sold a ([A-z]+) to ([A-z]+)");
        doc->get_Range()->Replace(regex, u"$3 bought a $2 from $1", options);

        ASSERT_EQ(useSubstitutions ? String(u"Paul bought a car from John.\rJoe bought a house from Jane.")
                                   : String(u"$3 bought a $2 from $1.\r$3 bought a $2 from $1."),
                  doc->GetText().Trim());
        //ExEnd
    }

    //ExStart
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:IReplacingCallback
    //ExFor:ReplaceAction
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExFor:ReplacingArgs.MatchNode
    //ExSummary:Shows how to insert an entire document's contents as a replacement of a match in a find-and-replace operation.
    void InsertDocumentAtReplace()
    {
        auto mainDoc = MakeObject<Document>(MyDir + u"Document insertion destination.docx");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        auto options = MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(MakeObject<ExRange::InsertDocumentAtReplaceHandler>());

        mainDoc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"\\[MY_DOCUMENT\\]"), u"", options);
        mainDoc->Save(ArtifactsDir + u"InsertDocument.InsertDocumentAtReplace.docx");

        TestInsertDocumentAtReplace(MakeObject<Document>(ArtifactsDir + u"InsertDocument.InsertDocumentAtReplace.docx"));
        //ExSkip
    }

    class InsertDocumentAtReplaceHandler : public IReplacingCallback
    {
    private:
        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            auto subDoc = MakeObject<Document>(MyDir + u"Document.docx");

            // Insert a document after the paragraph containing the matched text.
            auto para = System::ExplicitCast<Paragraph>(args->get_MatchNode()->get_ParentNode());
            InsertDocument(para, subDoc);

            // Remove the paragraph with the matched text.
            para->Remove();

            return ReplaceAction::Skip;
        }
    };

    /// <summary>
    /// Inserts all the nodes of another document after a paragraph or table.
    /// </summary>
    static void InsertDocument(SharedPtr<Node> insertionDestination, SharedPtr<Document> docToInsert)
    {
        if (insertionDestination->get_NodeType() == NodeType::Paragraph || insertionDestination->get_NodeType() == NodeType::Table)
        {
            SharedPtr<CompositeNode> dstStory = insertionDestination->get_ParentNode();

            auto importer = MakeObject<NodeImporter>(docToInsert, insertionDestination->get_Document(), ImportFormatMode::KeepSourceFormatting);

            for (const auto& srcSection : System::IterateOver(docToInsert->get_Sections()->LINQ_OfType<SharedPtr<Section>>()))
            {
                for (const auto& srcNode : System::IterateOver(srcSection->get_Body()))
                {
                    // Skip the node if it is the last empty paragraph in a section.
                    if (srcNode->get_NodeType() == NodeType::Paragraph)
                    {
                        auto para = System::ExplicitCast<Paragraph>(srcNode);
                        if (para->get_IsEndOfSection() && !para->get_HasChildNodes())
                        {
                            continue;
                        }
                    }

                    SharedPtr<Node> newNode = importer->ImportNode(srcNode, true);

                    dstStory->InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
        else
        {
            throw System::ArgumentException(u"The destination node must be either a paragraph or table.");
        }
    }
    //ExEnd

    static void TestInsertDocumentAtReplace(SharedPtr<Document> doc)
    {
        ASSERT_EQ(String(u"1) At text that can be identified by regex:\rHello World!\r") +
                      u"2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" + u"3) At a bookmark:",
                  doc->get_FirstSection()->get_Body()->GetText().Trim());
    }

    //ExStart
    //ExFor:FindReplaceOptions.Direction
    //ExFor:FindReplaceDirection
    //ExSummary:Shows how to determine which direction a find-and-replace operation traverses the document in.
    void Direction(FindReplaceDirection findReplaceDirection)
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert three runs which we can search for using a regex pattern.
        // Place one of those runs inside a text box.
        builder->Writeln(u"Match 1.");
        builder->Writeln(u"Match 2.");
        builder->Writeln(u"Match 3.");
        builder->Writeln(u"Match 4.");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        auto options = MakeObject<FindReplaceOptions>();

        // Assign a custom callback to the "ReplacingCallback" property.
        auto callback = MakeObject<ExRange::TextReplacementRecorder>();
        options->set_ReplacingCallback(callback);

        // Set the "Direction" property to "FindReplaceDirection.Backward" to get the find-and-replace
        // operation to start from the end of the range, and traverse back to the beginning.
        // Set the "Direction" property to "FindReplaceDirection.Backward" to get the find-and-replace
        // operation to start from the beginning of the range, and traverse to the end.
        options->set_Direction(findReplaceDirection);

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"Match \\d*"), u"Replacement", options);

        ASSERT_EQ(String(u"Replacement.\r") + u"Replacement.\r" + u"Replacement.\r" + u"Replacement.", doc->GetText().Trim());

        switch (findReplaceDirection)
        {
        case FindReplaceDirection::Forward:
            ASPOSE_ASSERT_EQ(MakeArray<String>({u"Match 1", u"Match 2", u"Match 3", u"Match 4"}), callback->get_Matches());
            break;

        case FindReplaceDirection::Backward:
            ASPOSE_ASSERT_EQ(MakeArray<String>({u"Match 4", u"Match 3", u"Match 2", u"Match 1"}), callback->get_Matches());
            break;
        }
    }

    /// <summary>
    /// Records all matches that occur during a find-and-replace operation in the order that they take place.
    /// </summary>
    class TextReplacementRecorder : public IReplacingCallback
    {
    public:
        SharedPtr<System::Collections::Generic::List<String>> get_Matches()
        {
            return mMatches;
        }

        TextReplacementRecorder() : mMatches(MakeObject<System::Collections::Generic::List<String>>())
        {
        }

    private:
        SharedPtr<System::Collections::Generic::List<String>> mMatches;

        ReplaceAction Replacing(SharedPtr<ReplacingArgs> e) override
        {
            get_Matches()->Add(e->get_Match()->get_Value());
            return ReplaceAction::Replace;
        }
    };
    //ExEnd
};

} // namespace ApiExamples
