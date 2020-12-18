#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldDocProperty.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceDirection.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Importing/NodeImporter.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <drawing/color.h>
#include <system/collections/ienumerable.h>
#include <system/collections/list.h>
#include <system/convert.h>
#include <system/date_time.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/test_tools.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Replacing;

namespace ApiExamples {

class ExRange : public ApiExampleBase
{
public:
    void ReplaceSimple()
    {
        //ExStart
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExFor:FindReplaceOptions
        //ExFor:FindReplaceOptions.MatchCase
        //ExFor:FindReplaceOptions.FindWholeWordsOnly
        //ExSummary:Simple find and replace operation.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello _CustomerName_,");

        // Check the document contains what we are about to test
        std::cout << doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->GetText() << std::endl;

        auto options = MakeObject<FindReplaceOptions>();
        options->set_MatchCase(false);
        options->set_FindWholeWordsOnly(false);

        doc->get_Range()->Replace(u"_CustomerName_", u"James Bond", options);

        doc->Save(ArtifactsDir + u"Range.ReplaceSimple.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Range.ReplaceSimple.docx");

        ASSERT_EQ(u"Hello James Bond,", doc->GetText().Trim());
    }

    void IgnoreDeleted(bool isIgnoreDeleted)
    {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreDeleted
        //ExSummary:Shows how to ignore text inside delete revisions.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert non-revised text
        builder->Writeln(u"Deleted");
        builder->Write(u"Text");

        // Remove first paragraph with tracking revisions
        doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->Remove();
        doc->StopTrackRevisions();

        auto regex = MakeObject<System::Text::RegularExpressions::Regex>(u"e");
        auto options = MakeObject<FindReplaceOptions>();

        // Replace 'e' in document while ignoring/not ignoring deleted text
        options->set_IgnoreDeleted(isIgnoreDeleted);
        doc->get_Range()->Replace(regex, u"*", options);

        ASSERT_EQ(doc->GetText().Trim(), isIgnoreDeleted ? String(u"Deleted\rT*xt") : String(u"D*l*t*d\rT*xt"));
        //ExEnd
    }

    void IgnoreInserted(bool isIgnoreInserted)
    {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreInserted
        //ExSummary:Shows how to ignore text inside insert revisions.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert text with tracking revisions
        doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
        builder->Writeln(u"Inserted");
        doc->StopTrackRevisions();

        // Insert non-revised text
        builder->Write(u"Text");

        auto regex = MakeObject<System::Text::RegularExpressions::Regex>(u"e");
        auto options = MakeObject<FindReplaceOptions>();

        // Replace 'e' in document while ignoring/not ignoring inserted text
        options->set_IgnoreInserted(isIgnoreInserted);
        doc->get_Range()->Replace(regex, u"*", options);

        ASSERT_EQ(doc->GetText().Trim(), isIgnoreInserted ? String(u"Inserted\rT*xt") : String(u"Ins*rt*d\rT*xt"));
        //ExEnd
    }

    void IgnoreFields(bool isIgnoreFields)
    {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreFields
        //ExSummary:Shows how to ignore text inside fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert field with text inside
        builder->InsertField(u"INCLUDETEXT", u"Text in field");

        auto regex = MakeObject<System::Text::RegularExpressions::Regex>(u"e");
        auto options = MakeObject<FindReplaceOptions>();
        // Replace 'e' in document ignoring/not ignoring text inside field
        options->set_IgnoreFields(isIgnoreFields);

        doc->get_Range()->Replace(regex, u"*", options);

        ASSERT_EQ(doc->GetText(), isIgnoreFields ? String(u"\u0013INCLUDETEXT\u0014Text in field\u0015\f")
                                                 : String(u"\u0013INCLUDETEXT\u0014T*xt in fi*ld\u0015\f"));
        //ExEnd
    }

    void UpdateFieldsInRange()
    {
        //ExStart
        //ExFor:Range.UpdateFields
        //ExSummary:Shows how to update document fields in the body of the first section only.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a field that will display the value in the document's body text
        auto field = System::DynamicCast<FieldDocProperty>(builder->InsertField(u" DOCPROPERTY Category"));

        // Set the value of the property that should be displayed by the field
        doc->get_BuiltInDocumentProperties()->set_Category(u"MyCategory");

        // Some field types need to be explicitly updated before they can display their expected values
        ASSERT_EQ(String::Empty, field->get_Result());

        // Update all the fields in the first section of the document, which includes the field we just inserted
        doc->get_FirstSection()->get_Range()->UpdateFields();

        ASSERT_EQ(u"MyCategory", field->get_Result());
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
        //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace all occurrences of words "sad" or "mad" to "bad".
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"sad mad bad");

        ASSERT_EQ(u"sad mad bad", doc->GetText().Trim());

        auto options = MakeObject<FindReplaceOptions>();
        options->set_MatchCase(false);
        options->set_FindWholeWordsOnly(false);

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"[s|m]ad"), u"bad", options);

        ASSERT_EQ(u"bad bad bad", doc->GetText().Trim());
        //ExEnd
    }

    //ExStart
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:ReplacingArgs.Replacement
    //ExFor:IReplacingCallback
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExSummary:Replaces text specified with regular expression with HTML.
    void ReplaceWithInsertHtml()
    {
        // Open the document
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello <CustomerName>,");

        auto options = MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(MakeObject<ExRange::ReplaceWithHtmlEvaluator>());

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u" <CustomerName>,"), String::Empty, options);

        // Save the modified document
        doc->Save(ArtifactsDir + u"Range.ReplaceWithInsertHtml.docx");
        ASSERT_EQ(u"James Bond, Hello\r\x000c", MakeObject<Document>(ArtifactsDir + u"Range.ReplaceWithInsertHtml.docx")->GetText());
        //ExSkip
    }

private:
    class ReplaceWithHtmlEvaluator : public IReplacingCallback
    {
    public:
        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            auto builder = MakeObject<DocumentBuilder>(System::DynamicCast<Document>(args->get_MatchNode()->get_Document()));
            builder->MoveTo(args->get_MatchNode());

            // Replace '<CustomerName>' text with a red bold name
            builder->InsertHtml(u"<b><font color='red'>James Bond, </font></b>");
            args->set_Replacement(u"");

            return ReplaceAction::Replace;
        }
    };
    //ExEnd

public:
    //ExStart
    //ExFor:FindReplaceOptions.ApplyFont
    //ExFor:FindReplaceOptions.Direction
    //ExFor:FindReplaceOptions.ReplacingCallback
    //ExFor:ReplacingArgs.GroupIndex
    //ExFor:ReplacingArgs.GroupName
    //ExFor:ReplacingArgs.Match
    //ExFor:ReplacingArgs.MatchOffset
    //ExSummary:Shows how to apply a different font to new content via FindReplaceOptions.
    void ReplaceNumbersAsHex()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(String(u"Numbers that will be converted to hexadecimal and highlighted:\n") + u"123, 456, 789 and 17379.");

        auto options = MakeObject<FindReplaceOptions>();
        // Highlight newly inserted content with a color
        options->get_ApplyFont()->set_HighlightColor(System::Drawing::Color::get_LightGray());

        // Apply an IReplacingCallback to make the replacement to convert integers into hex equivalents,
        // and then to count replacements in the order they take place
        options->set_ReplacingCallback(MakeObject<ExRange::NumberHexer>());
        // By default, text is searched for replacements front to back, but we can change it to go the other way
        options->set_Direction(FindReplaceDirection::Backward);

        int count = doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"[0-9]+"), u"", options);

        ASSERT_EQ(4, count);
        ASSERT_EQ(String(u"Numbers that will be converted to hexadecimal and highlighted:\r") + u"0x7B, 0x1C8, 0x315 and 0x43E3.",
                  doc->GetText().Trim());

        auto isLightGray = [](SharedPtr<Run> r)
        {
            return r->get_Font()->get_HighlightColor().ToArgb() == System::Drawing::Color::get_LightGray().ToArgb();
        };

        ASSERT_EQ(4, doc->GetChildNodes(NodeType::Run, true)->LINQ_OfType<SharedPtr<Run>>()->LINQ_Count(isLightGray));
    }

private:
    /// <summary>
    /// Replaces Arabic numbers with hexadecimal equivalents and appends the number of each replacement.
    /// </summary>
    class NumberHexer : public IReplacingCallback
    {
    public:
        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            mCurrentReplacementNumber++;

            // Parse numbers
            int number = System::Convert::ToInt32(args->get_Match()->get_Value());

            // And write it as HEX
            args->set_Replacement(String::Format(u"0x{0:X}", number));

            std::cout << "Match #" << mCurrentReplacementNumber << std::endl;
            std::cout << "\tOriginal value:\t" << args->get_Match()->get_Value() << std::endl;
            std::cout << "\tReplacement:\t" << args->get_Replacement() << std::endl;
            std::cout << String::Format(u"\tOffset in parent {0} node:\t{1}", args->get_MatchNode()->get_NodeType(), args->get_MatchOffset())
                      << std::endl;

            std::cout << (String::IsNullOrEmpty(args->get_GroupName()) ? String::Format(u"\tGroup index:\t{0}", args->get_GroupIndex())
                                                                               : String::Format(u"\tGroup name:\t{0}", args->get_GroupName()))
                      << std::endl;

            return ReplaceAction::Replace;
        }

        NumberHexer() : mCurrentReplacementNumber(0)
        {
        }

    private:
        int mCurrentReplacementNumber;
    };
    //ExEnd

public:
    void ApplyParagraphFormat()
    {
        //ExStart
        //ExFor:FindReplaceOptions.ApplyParagraphFormat
        //ExSummary:Shows how to affect the format of paragraphs with successful replacements.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Every paragraph that ends with a full stop like this one will be right aligned.");
        builder->Writeln(u"This one will not!");
        builder->Writeln(u"And this one will.");

        auto options = MakeObject<FindReplaceOptions>();
        options->get_ApplyParagraphFormat()->set_Alignment(ParagraphAlignment::Right);

        int count = doc->get_Range()->Replace(u".&p", u"!&p", options);
        ASSERT_EQ(2, count);

        doc->Save(ArtifactsDir + u"Range.ApplyParagraphFormat.docx");
        //ExEnd

        auto savedDoc = MakeObject<Document>(ArtifactsDir + u"Range.ApplyParagraphFormat.docx");
        SharedPtr<ParagraphCollection> paragraphs = savedDoc->get_FirstSection()->get_Body()->get_Paragraphs();

        ASSERT_EQ(ParagraphAlignment::Right, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Alignment());
        ASSERT_EQ(ParagraphAlignment::Left, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Alignment());
        ASSERT_EQ(ParagraphAlignment::Right, paragraphs->idx_get(2)->get_ParagraphFormat()->get_Alignment());
    }

    void DeleteSelection()
    {
        //ExStart
        //ExFor:Node.Range
        //ExFor:Range.Delete
        //ExSummary:Shows how to delete all characters of a range.
        // Insert two sections into a blank document
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Section 1. ");
        builder->InsertBreak(BreakType::SectionBreakContinuous);
        builder->Write(u"Section 2.");

        // Verify the whole text of the document
        ASSERT_EQ(u"Section 1. \fSection 2.", doc->GetText().Trim());

        // Delete the first section from the document
        doc->get_Sections()->idx_get(0)->get_Range()->Delete();

        // Check the first section was deleted by looking at the text of the whole document again
        ASSERT_EQ(u"Section 2.", doc->GetText().Trim());
        //ExEnd
    }

    void RangesGetText()
    {
        //ExStart
        //ExFor:Range
        //ExFor:Range.Text
        //ExSummary:Shows how to get plain, unformatted text of a range.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Hello world!");

        ASSERT_EQ(u"Hello world!", doc->get_Range()->get_Text().Trim());
        //ExEnd
    }

    void UseLegacyOrder(bool isUseLegacyOrder)
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert three runs which can be used as tags, with the second placed inside a text box.
        builder->Writeln(u"[tag 1]");
        SharedPtr<Shape> textBox = builder->InsertShape(ShapeType::TextBox, 100, 50);
        builder->Writeln(u"[tag 3]");
        builder->MoveTo(textBox->get_FirstParagraph());
        builder->Write(u"[tag 2]");

        auto options = MakeObject<FindReplaceOptions>();
        auto callback = MakeObject<ExRange::TextReplacementTracker>();
        options->set_ReplacingCallback(callback);

        // When a text replacement is performed, all of the runs of a document have their contents searched
        // for every instance of the string that we wish to replace.
        // This flag can change the search priority of runs inside text boxes.
        options->set_UseLegacyOrder(isUseLegacyOrder);

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"\\[tag \\d*\\]"), u"", options);

        // Using legacy order goes through all runs of a range in sequential order.
        // Not using legacy order goes through runs within text boxes after all runs outside of text boxes have been searched.
        SharedPtr<System::Collections::Generic::List<String>> matches = isUseLegacyOrder ? [&]{ String init_0[] = {u"[tag 1]", u"[tag 2]", u"[tag 3]"}; auto list_0 = MakeObject<System::Collections::Generic::List<String>>(); list_0->AddInitializer(3, init_0); return list_0; }() : [&]{ String init_1[] = {u"[tag 1]", u"[tag 3]", u"[tag 2]"}; auto list_1 = MakeObject<System::Collections::Generic::List<String>>(); list_1->AddInitializer(3, init_1); return list_1; }();
        ASPOSE_ASSERT_EQ(matches, callback->get_Matches());
    }

private:
    /// <summary>
    /// Creates a list of string matches from a regex-based text find-and-replacement operation
    /// in the order that they are encountered.
    /// </summary>
    class TextReplacementTracker : public IReplacingCallback
    {
    public:
        SharedPtr<System::Collections::Generic::List<String>> get_Matches()
        {
            return mMatches;
        }

        ReplaceAction Replacing(SharedPtr<ReplacingArgs> e) override
        {
            get_Matches()->Add(e->get_Match()->get_Value());
            return ReplaceAction::Replace;
        }

        TextReplacementTracker() : mMatches(MakeObject<System::Collections::Generic::List<String>>())
        {
        }

    private:
        SharedPtr<System::Collections::Generic::List<String>> mMatches;
    };
    //ExEnd

public:
    void UseSubstitutions()
    {
        //ExStart
        //ExFor:FindReplaceOptions.UseSubstitutions
        //ExSummary:Shows how to replace text with substitutions.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"John sold a car to Paul.");
        builder->Writeln(u"Jane sold a house to Joe.");

        // Perform a find-and-replace operation on a range's text contents
        // while preserving some elements from the replaced text using substitutions.
        auto options = MakeObject<FindReplaceOptions>();
        options->set_UseSubstitutions(true);

        auto regex = MakeObject<System::Text::RegularExpressions::Regex>(u"([A-z]+) sold a ([A-z]+) to ([A-z]+)");
        doc->get_Range()->Replace(regex, u"$3 bought a $2 from $1", options);

        ASSERT_EQ(doc->GetText(), u"Paul bought a car from John.\rJoe bought a house from Jane.\r\f");
        //ExEnd
    }

    //ExStart
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:IReplacingCallback
    //ExFor:ReplaceAction
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExFor:ReplacingArgs.MatchNode
    //ExFor:FindReplaceDirection
    //ExSummary:Shows how to insert content of one document into another during a customized find and replace operation.
    void InsertDocumentAtReplace()
    {
        auto mainDoc = MakeObject<Document>(MyDir + u"Document insertion destination.docx");

        auto options = MakeObject<FindReplaceOptions>();
        options->set_Direction(FindReplaceDirection::Backward);
        options->set_ReplacingCallback(MakeObject<ExRange::InsertDocumentAtReplaceHandler>());

        mainDoc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"\\[MY_DOCUMENT\\]"), u"", options);
        mainDoc->Save(ArtifactsDir + u"InsertDocument.InsertDocumentAtReplace.docx");

        TestInsertDocumentAtReplace(MakeObject<Document>(ArtifactsDir + u"InsertDocument.InsertDocumentAtReplace.docx"));
        //ExSkip
    }

private:
    class InsertDocumentAtReplaceHandler : public IReplacingCallback
    {
    public:
        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            auto subDoc = MakeObject<Document>(MyDir + u"Document.docx");

            // Insert a document after the paragraph, containing the match text
            auto para = System::DynamicCast<Paragraph>(args->get_MatchNode()->get_ParentNode());
            InsertDocument(para, subDoc);

            // Remove the paragraph with the match text
            para->Remove();

            return ReplaceAction::Skip;
        }
    };

    /// <summary>
    /// Inserts content of the external document after the specified node.
    /// </summary>
    static void InsertDocument(SharedPtr<Node> insertionDestination, SharedPtr<Document> docToInsert)
    {
        // Make sure that the node is either a paragraph or table
        if (insertionDestination->get_NodeType() == NodeType::Paragraph || insertionDestination->get_NodeType() == NodeType::Table)
        {
            // We will be inserting into the parent of the destination paragraph
            SharedPtr<CompositeNode> dstStory = insertionDestination->get_ParentNode();

            // This object will be translating styles and lists during the import
            auto importer = MakeObject<NodeImporter>(docToInsert, insertionDestination->get_Document(), ImportFormatMode::KeepSourceFormatting);

            // Loop through all block level nodes in the body of the section
            for (auto srcSection : System::IterateOver(docToInsert->get_Sections()->LINQ_OfType<SharedPtr<Section>>()))
            {
                for (auto srcNode : System::IterateOver(srcSection->get_Body()))
                {
                    // Skip the node if it is a last empty paragraph in a section
                    if (srcNode->get_NodeType() == NodeType::Paragraph)
                    {
                        auto para = System::DynamicCast<Paragraph>(srcNode);
                        if (para->get_IsEndOfSection() && !para->get_HasChildNodes())
                        {
                            continue;
                        }
                    }

                    // This creates a clone of the node, suitable for insertion into the destination document
                    SharedPtr<Node> newNode = importer->ImportNode(srcNode, true);

                    // Insert new node after the reference node
                    dstStory->InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
        else
        {
            throw System::ArgumentException(u"The destination node should be either a paragraph or table.");
        }
    }
    //ExEnd

protected:
    void TestInsertDocumentAtReplace(SharedPtr<Document> doc)
    {
        ASSERT_EQ(String(u"1) At text that can be identified by regex:\rHello World!\r") +
                      u"2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" + u"3) At a bookmark:",
                  doc->get_FirstSection()->get_Body()->GetText().Trim());
    }
};

} // namespace ApiExamples
