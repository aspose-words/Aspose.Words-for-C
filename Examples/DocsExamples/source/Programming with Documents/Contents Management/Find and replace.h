#pragma once

#include <cstdint>
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
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Replacing/FindReplaceDirection.h>
#include <Aspose.Words.Cpp/Replacing/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Replacing/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Replacing/ReplaceAction.h>
#include <Aspose.Words.Cpp/Replacing/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <drawing/color.h>
#include <system/collections/list.h>
#include <system/convert.h>
#include <system/date_time.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/object_ext.h>
#include <system/text/regularexpressions/group.h>
#include <system/text/regularexpressions/group_collection.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/regex_options.h>
#include <system/text/string_builder.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Management {

class FindAndReplace : public DocsExamplesBase
{
public:
    void SimpleFindReplace()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello _CustomerName_,");
        std::cout << (String(u"Original document text: ") + doc->get_Range()->get_Text()) << std::endl;

        doc->get_Range()->Replace(u"_CustomerName_", u"James Bond", MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));

        std::cout << (String(u"Document text after replace: ") + doc->get_Range()->get_Text()) << std::endl;

        // Save the modified document
        doc->Save(ArtifactsDir + u"FindAndReplace.SimpleFindReplace.docx");
    }

    void FindAndHighlight()
    {
        //ExStart:FindAndHighlight
        auto doc = MakeObject<Document>(MyDir + u"Find and highlight.docx");

        auto options = MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(MakeObject<FindAndReplace::ReplaceEvaluatorFindAndHighlight>());
        options->set_Direction(FindReplaceDirection::Backward);

        auto regex = MakeObject<System::Text::RegularExpressions::Regex>(u"your document", System::Text::RegularExpressions::RegexOptions::IgnoreCase);
        doc->get_Range()->Replace(regex, u"", options);

        doc->Save(ArtifactsDir + u"FindAndReplace.FindAndHighlight.docx");
        //ExEnd:FindAndHighlight
    }

    //ExStart:ReplaceEvaluatorFindAndHighlight
    class ReplaceEvaluatorFindAndHighlight : public IReplacingCallback
    {
    private:
        /// <summary>
        /// This method is called by the Aspose.Words find and replace engine for each match.
        /// This method highlights the match string, even if it spans multiple runs.
        /// </summary>
        ReplaceAction Replacing(SharedPtr<ReplacingArgs> e) override
        {
            // This is a Run node that contains either the beginning or the complete match.
            SharedPtr<Node> currentNode = e->get_MatchNode();

            // The first (and may be the only) run can contain text before the match,
            // in this case it is necessary to split the run.
            if (e->get_MatchOffset() > 0)
            {
                currentNode = SplitRun(System::ExplicitCast<Aspose::Words::Run>(currentNode), e->get_MatchOffset());
            }

            // This array is used to store all nodes of the match for further highlighting.
            SharedPtr<System::Collections::Generic::List<SharedPtr<Run>>> runs = MakeObject<System::Collections::Generic::List<SharedPtr<Run>>>();

            // Find all runs that contain parts of the match string.
            int remainingLength = e->get_Match()->get_Value().get_Length();
            while (remainingLength > 0 && currentNode != nullptr && currentNode->GetText().get_Length() <= remainingLength)
            {
                runs->Add(System::ExplicitCast<Aspose::Words::Run>(currentNode));
                remainingLength -= currentNode->GetText().get_Length();

                // Select the next Run node.
                // Have to loop because there could be other nodes such as BookmarkStart etc.
                do
                {
                    currentNode = currentNode->get_NextSibling();
                } while (currentNode != nullptr && currentNode->get_NodeType() != NodeType::Run);
            }

            // Split the last run that contains the match if there is any text left.
            if (currentNode != nullptr && remainingLength > 0)
            {
                SplitRun(System::ExplicitCast<Aspose::Words::Run>(currentNode), remainingLength);
                runs->Add(System::ExplicitCast<Aspose::Words::Run>(currentNode));
            }

            // Now highlight all runs in the sequence.
            for (const auto& run : runs)
            {
                run->get_Font()->set_HighlightColor(System::Drawing::Color::get_Yellow());
            }

            // Signal to the replace engine to do nothing because we have already done all what we wanted.
            return ReplaceAction::Skip;
        }
    };
    //ExEnd:ReplaceEvaluatorFindAndHighlight

    //ExStart:SplitRun
    /// <summary>
    /// Splits text of the specified run into two runs.
    /// Inserts the new run just after the specified run.
    /// </summary>
    static SharedPtr<Run> SplitRun(SharedPtr<Run> run, int position)
    {
        auto afterRun = System::ExplicitCast<Aspose::Words::Run>(run->Clone(true));
        afterRun->set_Text(run->get_Text().Substring(position));

        run->set_Text(run->get_Text().Substring(0, position));
        run->get_ParentNode()->InsertAfter(afterRun, run);

        return afterRun;
    }
    //ExEnd:SplitRun

    void MetaCharactersInSearchPattern()
    {
        /* meta-characters
                    &p - paragraph break
                    &b - section break
                    &m - page break
                    &l - manual line break
                    */

        //ExStart:MetaCharactersInSearchPattern
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"This is Line 1");
        builder->Writeln(u"This is Line 2");

        doc->get_Range()->Replace(u"This is Line 1&pThis is Line 2", u"This is replaced line");

        builder->MoveToDocumentEnd();
        builder->Write(u"This is Line 1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"This is Line 2");

        doc->get_Range()->Replace(u"This is Line 1&mThis is Line 2", u"Page break is replaced with new text.");

        doc->Save(ArtifactsDir + u"FindAndReplace.MetaCharactersInSearchPattern.docx");
        //ExEnd:MetaCharactersInSearchPattern
    }

    void ReplaceTextContainingMetaCharacters()
    {
        //ExStart:ReplaceTextContainingMetaCharacters
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(u"First section");
        builder->Writeln(u"  1st paragraph");
        builder->Writeln(u"  2nd paragraph");
        builder->Writeln(u"{insert-section}");
        builder->Writeln(u"Second section");
        builder->Writeln(u"  1st paragraph");

        auto findReplaceOptions = MakeObject<FindReplaceOptions>();
        findReplaceOptions->get_ApplyParagraphFormat()->set_Alignment(ParagraphAlignment::Center);

        // Double each paragraph break after word "section", add kind of underline and make it centered.
        int count = doc->get_Range()->Replace(u"section&p", u"section&p----------------------&p", findReplaceOptions);

        // Insert section break instead of custom text tag.
        count = doc->get_Range()->Replace(u"{insert-section}", u"&b", findReplaceOptions);

        doc->Save(ArtifactsDir + u"FindAndReplace.ReplaceTextContainingMetaCharacters.docx");
        //ExEnd:ReplaceTextContainingMetaCharacters
    }

    void IgnoreTextInsideFields()
    {
        //ExStart:IgnoreTextInsideFields
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert field with text inside.
        builder->InsertField(u"INCLUDETEXT", u"Text in field");

        auto options = MakeObject<FindReplaceOptions>();
        options->set_IgnoreFields(true);

        auto regex = MakeObject<System::Text::RegularExpressions::Regex>(u"e");
        doc->get_Range()->Replace(regex, u"*", options);

        std::cout << doc->GetText() << std::endl;

        options->set_IgnoreFields(false);
        doc->get_Range()->Replace(regex, u"*", options);

        std::cout << doc->GetText() << std::endl;
        //ExEnd:IgnoreTextInsideFields
    }

    void IgnoreTextInsideDeleteRevisions()
    {
        //ExStart:IgnoreTextInsideDeleteRevisions
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert non-revised text.
        builder->Writeln(u"Deleted");
        builder->Write(u"Text");

        // Remove first paragraph with tracking revisions.
        doc->StartTrackRevisions(u"author", System::DateTime::get_Now());
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->Remove();
        doc->StopTrackRevisions();

        auto options = MakeObject<FindReplaceOptions>();
        options->set_IgnoreDeleted(true);

        auto regex = MakeObject<System::Text::RegularExpressions::Regex>(u"e");
        doc->get_Range()->Replace(regex, u"*", options);

        std::cout << doc->GetText() << std::endl;

        options->set_IgnoreDeleted(false);
        doc->get_Range()->Replace(regex, u"*", options);

        std::cout << doc->GetText() << std::endl;
        //ExEnd:IgnoreTextInsideDeleteRevisions
    }

    void IgnoreTextInsideInsertRevisions()
    {
        //ExStart:IgnoreTextInsideInsertRevisions
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert text with tracking revisions.
        doc->StartTrackRevisions(u"author", System::DateTime::get_Now());
        builder->Writeln(u"Inserted");
        doc->StopTrackRevisions();

        // Insert non-revised text.
        builder->Write(u"Text");

        auto options = MakeObject<FindReplaceOptions>();
        options->set_IgnoreInserted(true);

        auto regex = MakeObject<System::Text::RegularExpressions::Regex>(u"e");
        doc->get_Range()->Replace(regex, u"*", options);

        std::cout << doc->GetText() << std::endl;

        options->set_IgnoreInserted(false);
        doc->get_Range()->Replace(regex, u"*", options);

        std::cout << doc->GetText() << std::endl;
        //ExEnd:IgnoreTextInsideInsertRevisions
    }

    void ReplaceHtmlTextWithMetaCharacters()
    {
        //ExStart:ReplaceHtmlTextWithMetaCharacters
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"{PLACEHOLDER}");

        auto findReplaceOptions = MakeObject<FindReplaceOptions>();
        findReplaceOptions->set_ReplacingCallback(MakeObject<FindAndReplace::FindAndInsertHtml>());

        doc->get_Range()->Replace(u"{PLACEHOLDER}", u"<p>&ldquo;Some Text&rdquo;</p>", findReplaceOptions);

        doc->Save(ArtifactsDir + u"FindAndReplace.ReplaceHtmlTextWithMetaCharacters.docx");
        //ExEnd:ReplaceHtmlTextWithMetaCharacters
    }

    //ExStart:ReplaceHtmlFindAndInsertHtml
    class FindAndInsertHtml final : public IReplacingCallback
    {
    private:
        ReplaceAction Replacing(SharedPtr<ReplacingArgs> e) override
        {
            SharedPtr<Node> currentNode = e->get_MatchNode();

            auto builder = MakeObject<DocumentBuilder>(System::AsCast<Document>(e->get_MatchNode()->get_Document()));
            builder->MoveTo(currentNode);
            builder->InsertHtml(e->get_Replacement());

            currentNode->Remove();

            return ReplaceAction::Skip;
        }
    };
    //ExEnd:ReplaceHtmlFindAndInsertHtml

    void ReplaceTextInFooter()
    {
        //ExStart:ReplaceTextInFooter
        auto doc = MakeObject<Document>(MyDir + u"Footer.docx");

        SharedPtr<HeaderFooterCollection> headersFooters = doc->get_FirstSection()->get_HeadersFooters();
        SharedPtr<HeaderFooter> footer = headersFooters->idx_get(HeaderFooterType::FooterPrimary);

        auto options = MakeObject<FindReplaceOptions>();
        options->set_MatchCase(false);
        options->set_FindWholeWordsOnly(false);

        footer->get_Range()->Replace(u"(C) 2006 Aspose Pty Ltd.", u"Copyright (C) 2020 by Aspose Pty Ltd.", options);

        doc->Save(ArtifactsDir + u"FindAndReplace.ReplaceTextInFooter.docx");
        //ExEnd:ReplaceTextInFooter
    }

    //ExStart:ShowChangesForHeaderAndFooterOrders
    void ShowChangesForHeaderAndFooterOrders()
    {
        auto logger = MakeObject<FindAndReplace::ReplaceLog>();

        auto doc = MakeObject<Document>(MyDir + u"Footer.docx");
        SharedPtr<Section> firstPageSection = doc->get_FirstSection();

        auto options = MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(logger);

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"(header|footer)"), u"", options);

        doc->Save(ArtifactsDir + u"FindAndReplace.ShowChangesForHeaderAndFooterOrders.docx");

        logger->ClearText();

        firstPageSection->get_PageSetup()->set_DifferentFirstPageHeaderFooter(false);

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"(header|footer)"), u"", options);
    }

    class ReplaceLog : public IReplacingCallback
    {
    public:
        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            mTextBuilder->AppendLine(args->get_MatchNode()->GetText());
            return ReplaceAction::Skip;
        }

        void ClearText()
        {
            mTextBuilder->Clear();
        }

        ReplaceLog() : mTextBuilder(MakeObject<System::Text::StringBuilder>())
        {
        }

    private:
        SharedPtr<System::Text::StringBuilder> mTextBuilder;
    };
    //ExEnd:ShowChangesForHeaderAndFooterOrders

    void ReplaceTextWithField()
    {
        auto doc = MakeObject<Document>(MyDir + u"Replace text with fields.docx");

        auto options = MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(MakeObject<FindAndReplace::ReplaceTextWithFieldHandler>(FieldType::FieldMergeField));

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"PlaceHolder(\\d+)"), u"", options);

        doc->Save(ArtifactsDir + u"FindAndReplace.ReplaceTextWithField.docx");
    }

    class ReplaceTextWithFieldHandler : public IReplacingCallback
    {
    public:
        ReplaceTextWithFieldHandler(FieldType type) : mFieldType(((Aspose::Words::Fields::FieldType)0))
        {
            mFieldType = type;
        }

        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            SharedPtr<System::Collections::Generic::List<SharedPtr<Run>>> runs = FindAndSplitMatchRuns(args);

            auto builder = MakeObject<DocumentBuilder>(System::ExplicitCast<Document>(args->get_MatchNode()->get_Document()));
            builder->MoveTo(runs->idx_get(runs->get_Count() - 1));

            // Calculate the field's name from the FieldType enumeration by removing
            // the first instance of "Field" from the text. This works for almost all of the field types.
            String fieldName = System::ObjectExt::ToString(mFieldType).ToUpper().Substring(5);

            // Insert the field into the document using the specified field type and the matched text as the field name.
            // If the fields you are inserting do not require this extra parameter, it can be removed from the string below.
            builder->InsertField(String::Format(u"{0} {1}", fieldName, args->get_Match()->get_Groups()->idx_get(0)));

            for (const auto& run : runs)
            {
                run->Remove();
            }

            return ReplaceAction::Skip;
        }

        /// <summary>
        /// Finds and splits the match runs and returns them in an List.
        /// </summary>
        SharedPtr<System::Collections::Generic::List<SharedPtr<Run>>> FindAndSplitMatchRuns(SharedPtr<ReplacingArgs> args)
        {
            // This is a Run node that contains either the beginning or the complete match.
            SharedPtr<Node> currentNode = args->get_MatchNode();

            // The first (and may be the only) run can contain text before the match,
            // In this case it is necessary to split the run.
            if (args->get_MatchOffset() > 0)
            {
                currentNode = SplitRun(System::ExplicitCast<Aspose::Words::Run>(currentNode), args->get_MatchOffset());
            }

            // This array is used to store all nodes of the match for further removing.
            SharedPtr<System::Collections::Generic::List<SharedPtr<Run>>> runs = MakeObject<System::Collections::Generic::List<SharedPtr<Run>>>();

            // Find all runs that contain parts of the match string.
            int remainingLength = args->get_Match()->get_Value().get_Length();
            while (remainingLength > 0 && currentNode != nullptr && currentNode->GetText().get_Length() <= remainingLength)
            {
                runs->Add(System::ExplicitCast<Aspose::Words::Run>(currentNode));
                remainingLength -= currentNode->GetText().get_Length();

                do
                {
                    currentNode = currentNode->get_NextSibling();
                } while (currentNode != nullptr && currentNode->get_NodeType() != NodeType::Run);
            }

            // Split the last run that contains the match if there is any text left.
            if (currentNode != nullptr && remainingLength > 0)
            {
                SplitRun(System::ExplicitCast<Aspose::Words::Run>(currentNode), remainingLength);
                runs->Add(System::ExplicitCast<Aspose::Words::Run>(currentNode));
            }

            return runs;
        }

    private:
        FieldType mFieldType;

        /// <summary>
        /// Splits text of the specified run into two runs.
        /// Inserts the new run just after the specified run.
        /// </summary>
        SharedPtr<Run> SplitRun(SharedPtr<Run> run, int position)
        {
            auto afterRun = System::ExplicitCast<Aspose::Words::Run>(run->Clone(true));

            afterRun->set_Text(run->get_Text().Substring(position));
            run->set_Text(run->get_Text().Substring(0, position));

            run->get_ParentNode()->InsertAfter(afterRun, run);

            return afterRun;
        }
    };

    void ReplaceWithEvaluator()
    {
        //ExStart:ReplaceWithEvaluator
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"sad mad bad");

        auto options = MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(MakeObject<FindAndReplace::MyReplaceEvaluator>());

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"[s|m]ad"), u"", options);

        doc->Save(ArtifactsDir + u"FindAndReplace.ReplaceWithEvaluator.docx");
        //ExEnd:ReplaceWithEvaluator
    }

    //ExStart:MyReplaceEvaluator
    class MyReplaceEvaluator : public IReplacingCallback
    {
    public:
        MyReplaceEvaluator() : mMatchNumber(0)
        {
        }

    private:
        int mMatchNumber;

        /// <summary>
        /// This is called during a replace operation each time a match is found.
        /// This method appends a number to the match string and returns it as a replacement string.
        /// </summary>
        ReplaceAction Replacing(SharedPtr<ReplacingArgs> e) override
        {
            e->set_Replacement(System::Convert::ToString(e->get_Match()) + System::Convert::ToString(mMatchNumber));
            mMatchNumber++;

            return ReplaceAction::Replace;
        }
    };
    //ExEnd:MyReplaceEvaluator

    //ExStart:ReplaceWithHtml
    void ReplaceWithHtml()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello <CustomerName>,");

        auto options = MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(MakeObject<FindAndReplace::ReplaceWithHtmlEvaluator>(options));

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u" <CustomerName>,"), String::Empty, options);

        doc->Save(ArtifactsDir + u"FindAndReplace.ReplaceWithHtml.docx");
    }

    class ReplaceWithHtmlEvaluator : public IReplacingCallback
    {
    public:
        ReplaceWithHtmlEvaluator(SharedPtr<FindReplaceOptions> options)
        {
            mOptions = options;
        }

    private:
        SharedPtr<FindReplaceOptions> mOptions;

        /// <summary>
        /// NOTE: This is a simplistic method that will only work well when the match
        /// starts at the beginning of a run.
        /// </summary>
        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            auto builder = MakeObject<DocumentBuilder>(System::ExplicitCast<Document>(args->get_MatchNode()->get_Document()));
            builder->MoveTo(args->get_MatchNode());

            // Replace '<CustomerName>' text with a red bold name.
            builder->InsertHtml(u"<b><font color='red'>James Bond, </font></b>");
            args->set_Replacement(u"");

            return ReplaceAction::Replace;
        }
    };
    //ExEnd:ReplaceWithHtml

    void ReplaceWithRegex()
    {
        //ExStart:ReplaceWithRegex
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"sad mad bad");

        auto options = MakeObject<FindReplaceOptions>();

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"[s|m]ad"), u"bad", options);

        doc->Save(ArtifactsDir + u"FindAndReplace.ReplaceWithRegex.docx");
        //ExEnd:ReplaceWithRegex
    }

    void RecognizeAndSubstitutionsWithinReplacementPatterns()
    {
        //ExStart:RecognizeAndSubstitutionsWithinReplacementPatterns
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Jason give money to Paul.");

        auto regex = MakeObject<System::Text::RegularExpressions::Regex>(u"([A-z]+) give money to ([A-z]+)");

        auto options = MakeObject<FindReplaceOptions>();
        options->set_UseSubstitutions(true);

        doc->get_Range()->Replace(regex, u"$2 take money from $1", options);
        //ExEnd:RecognizeAndSubstitutionsWithinReplacementPatterns
    }

    void ReplaceWithString()
    {
        //ExStart:ReplaceWithString
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"sad mad bad");

        doc->get_Range()->Replace(u"sad", u"bad", MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));

        doc->Save(ArtifactsDir + u"FindAndReplace.ReplaceWithString.docx");
        //ExEnd:ReplaceWithString
    }

    //ExStart:UsingLegacyOrder
    void UsingLegacyOrder()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"[tag 1]");
        SharedPtr<Shape> textBox = builder->InsertShape(ShapeType::TextBox, 100, 50);
        builder->Writeln(u"[tag 3]");

        builder->MoveTo(textBox->get_FirstParagraph());
        builder->Write(u"[tag 2]");

        auto options = MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(MakeObject<FindAndReplace::ReplacingCallback>());
        options->set_UseLegacyOrder(true);

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"\\[(.*?)\\]"), u"", options);

        doc->Save(ArtifactsDir + u"FindAndReplace.UsingLegacyOrder.docx");
    }

    class ReplacingCallback : public IReplacingCallback
    {
    private:
        ReplaceAction Replacing(SharedPtr<ReplacingArgs> e) override
        {
            std::cout << e->get_Match()->get_Value();
            return ReplaceAction::Replace;
        }
    };
    //ExEnd:UsingLegacyOrder

    void ReplaceTextInTable()
    {
        //ExStart:ReplaceText
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::ExplicitCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        table->get_Range()->Replace(u"Carrots", u"Eggs", MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));
        table->get_LastRow()->get_LastCell()->get_Range()->Replace(u"50", u"20", MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));

        doc->Save(ArtifactsDir + u"FindAndReplace.ReplaceTextInTable.docx");
        //ExEnd:ReplaceText
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Contents_Management
