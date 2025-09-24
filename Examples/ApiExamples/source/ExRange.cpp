// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRange.h"

#include <testing/test_predicates.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/date_time.h>
#include <system/convert.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/color.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Importing/NodeImporter.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Replacing;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3205497822u, ::Aspose::Words::ApiExamples::ExRange::TextFindAndReplacementLogger, ThisTypeBaseTypesInfo);

Aspose::Words::Replacing::ReplaceAction ExRange::TextFindAndReplacementLogger::Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args)
{
    mLog->AppendLine(System::String::Format(u"\"{0}\" converted to \"{1}\" ", args->get_Match()->get_Value(), args->get_Replacement()) + System::String::Format(u"{0} characters into a {1} node.", args->get_MatchOffset(), args->get_MatchNode()->get_NodeType()));
    
    args->set_Replacement(System::String::Format(u"(Old value:\"{0}\") {1}", args->get_Match()->get_Value(), args->get_Replacement()));
    return Aspose::Words::Replacing::ReplaceAction::Replace;
}

System::String ExRange::TextFindAndReplacementLogger::GetLog()
{
    return mLog->ToString();
}

ExRange::TextFindAndReplacementLogger::TextFindAndReplacementLogger()
    : mLog(System::MakeObject<System::Text::StringBuilder>())
{
}

RTTI_INFO_IMPL_HASH(3694841938u, ::Aspose::Words::ApiExamples::ExRange::NumberHexer, ThisTypeBaseTypesInfo);

Aspose::Words::Replacing::ReplaceAction ExRange::NumberHexer::Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args)
{
    mCurrentReplacementNumber++;
    
    int32_t number = System::Convert::ToInt32(args->get_Match()->get_Value());
    
    args->set_Replacement(System::String::Format(u"0x{0:X}", number));
    
    mLog->AppendLine(System::String::Format(u"Match #{0}", mCurrentReplacementNumber));
    mLog->AppendLine(System::String::Format(u"\tOriginal value:\t{0}", args->get_Match()->get_Value()));
    mLog->AppendLine(System::String::Format(u"\tReplacement:\t{0}", args->get_Replacement()));
    mLog->AppendLine(System::String::Format(u"\tOffset in parent {0} node:\t{1}", args->get_MatchNode()->get_NodeType(), args->get_MatchOffset()));
    
    mLog->AppendLine(System::String::IsNullOrEmpty(args->get_GroupName()) ? System::String::Format(u"\tGroup index:\t{0}", args->get_GroupIndex()) : System::String::Format(u"\tGroup name:\t{0}", args->get_GroupName()));
    
    return Aspose::Words::Replacing::ReplaceAction::Replace;
}

System::String ExRange::NumberHexer::GetLog()
{
    return mLog->ToString();
}

ExRange::NumberHexer::NumberHexer() : mCurrentReplacementNumber(0)
    , mLog(System::MakeObject<System::Text::StringBuilder>())
{
}

RTTI_INFO_IMPL_HASH(4225915840u, ::Aspose::Words::ApiExamples::ExRange::TextReplacementTracker, ThisTypeBaseTypesInfo);

System::SharedPtr<System::Collections::Generic::List<System::String>> ExRange::TextReplacementTracker::get_Matches() const
{
    return mMatches;
}

Aspose::Words::Replacing::ReplaceAction ExRange::TextReplacementTracker::Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> e)
{
    get_Matches()->Add(e->get_Match()->get_Value());
    return Aspose::Words::Replacing::ReplaceAction::Replace;
}

ExRange::TextReplacementTracker::TextReplacementTracker()
    : mMatches(System::MakeObject<System::Collections::Generic::List<System::String>>())
{
}

RTTI_INFO_IMPL_HASH(3299162058u, ::Aspose::Words::ApiExamples::ExRange::InsertDocumentAtReplaceHandler, ThisTypeBaseTypesInfo);

Aspose::Words::Replacing::ReplaceAction ExRange::InsertDocumentAtReplaceHandler::Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args)
{
    auto subDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    // Insert a document after the paragraph containing the matched text.
    auto para = System::ExplicitCast<Aspose::Words::Paragraph>(args->get_MatchNode()->get_ParentNode());
    InsertDocument(para, subDoc);
    
    // Remove the paragraph with the matched text.
    para->Remove();
    
    return Aspose::Words::Replacing::ReplaceAction::Skip;
}

RTTI_INFO_IMPL_HASH(3535072694u, ::Aspose::Words::ApiExamples::ExRange::TextReplacementRecorder, ThisTypeBaseTypesInfo);

System::SharedPtr<System::Collections::Generic::List<System::String>> ExRange::TextReplacementRecorder::get_Matches() const
{
    return mMatches;
}

Aspose::Words::Replacing::ReplaceAction ExRange::TextReplacementRecorder::Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> e)
{
    get_Matches()->Add(e->get_Match()->get_Value());
    return Aspose::Words::Replacing::ReplaceAction::Replace;
}

ExRange::TextReplacementRecorder::TextReplacementRecorder()
    : mMatches(System::MakeObject<System::Collections::Generic::List<System::String>>())
{
}

RTTI_INFO_IMPL_HASH(3388899529u, ::Aspose::Words::ApiExamples::ExRange::ReplacingCallback, ThisTypeBaseTypesInfo);

System::String ExRange::ReplacingCallback::get_StartNodeText() const
{
    return pr_StartNodeText;
}

void ExRange::ReplacingCallback::set_StartNodeText(System::String value)
{
    pr_StartNodeText = value;
}

System::String ExRange::ReplacingCallback::get_EndNodeText() const
{
    return pr_EndNodeText;
}

void ExRange::ReplacingCallback::set_EndNodeText(System::String value)
{
    pr_EndNodeText = value;
}

Aspose::Words::Replacing::ReplaceAction ExRange::ReplacingCallback::Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> e)
{
    set_StartNodeText(e->get_MatchNode()->GetText().Trim());
    set_EndNodeText(e->get_MatchEndNode()->GetText().Trim());
    
    return Aspose::Words::Replacing::ReplaceAction::Replace;
}


RTTI_INFO_IMPL_HASH(1627746669u, ::Aspose::Words::ApiExamples::ExRange, ThisTypeBaseTypesInfo);

void ExRange::InsertDocument(System::SharedPtr<Aspose::Words::Node> insertionDestination, System::SharedPtr<Aspose::Words::Document> docToInsert)
{
    if (insertionDestination->get_NodeType() == Aspose::Words::NodeType::Paragraph || insertionDestination->get_NodeType() == Aspose::Words::NodeType::Table)
    {
        System::SharedPtr<Aspose::Words::CompositeNode> dstStory = insertionDestination->get_ParentNode();
        
        auto importer = System::MakeObject<Aspose::Words::NodeImporter>(docToInsert, insertionDestination->get_Document(), Aspose::Words::ImportFormatMode::KeepSourceFormatting);
        
        for (auto&& srcSection : System::IterateOver(docToInsert->get_Sections()->LINQ_OfType<System::SharedPtr<Aspose::Words::Section> >()))
        {
            for (auto&& srcNode : System::IterateOver(srcSection->get_Body()))
            {
                // Skip the node if it is the last empty paragraph in a section.
                if (srcNode->get_NodeType() == Aspose::Words::NodeType::Paragraph)
                {
                    auto para = System::ExplicitCast<Aspose::Words::Paragraph>(srcNode);
                    if (para->get_IsEndOfSection() && !para->get_HasChildNodes())
                    {
                        continue;
                    }
                }
                
                System::SharedPtr<Aspose::Words::Node> newNode = importer->ImportNode(srcNode, true);
                
                dstStory->InsertAfter<System::SharedPtr<Aspose::Words::Node>>(newNode, insertionDestination);
                insertionDestination = newNode;
            }
        }
    }
    else
    {
        throw System::ArgumentException(u"The destination node must be either a paragraph or table.");
    }
}

void ExRange::TestInsertDocumentAtReplace(System::SharedPtr<Aspose::Words::Document> doc)
{
    ASSERT_EQ(System::String(u"1) At text that can be identified by regex:\rHello World!\r") + u"2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" + u"3) At a bookmark:", doc->get_FirstSection()->get_Body()->GetText().Trim());
}


namespace gtest_test
{

class ExRange : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExRange> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExRange>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExRange> ExRange::s_instance;

} // namespace gtest_test

void ExRange::Replace()
{
    //ExStart
    //ExFor:Range.Replace(String, String)
    //ExSummary:Shows how to perform a find-and-replace text operation on the contents of a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Greetings, _FullName_!");
    
    // Perform a find-and-replace operation on our document's contents and verify the number of replacements that took place.
    int32_t replacementCount = doc->get_Range()->Replace(u"_FullName_", u"John Doe");
    
    ASSERT_EQ(1, replacementCount);
    ASSERT_EQ(u"Greetings, John Doe!", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRange, Replace)
{
    s_instance->Replace();
}

} // namespace gtest_test

void ExRange::ReplaceMatchCase(bool matchCase)
{
    //ExStart
    //ExFor:Range.Replace(String, String, FindReplaceOptions)
    //ExFor:FindReplaceOptions
    //ExFor:FindReplaceOptions.MatchCase
    //ExSummary:Shows how to toggle case sensitivity when performing a find-and-replace operation.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Ruby bought a ruby necklace.");
    
    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    
    // Set the "MatchCase" flag to "true" to apply case sensitivity while finding strings to replace.
    // Set the "MatchCase" flag to "false" to ignore character case while searching for text to replace.
    options->set_MatchCase(matchCase);
    
    doc->get_Range()->Replace(u"Ruby", u"Jade", options);
    
    ASSERT_EQ(matchCase ? System::String(u"Jade bought a ruby necklace.") : System::String(u"Jade bought a Jade necklace."), doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

using ExRange_ReplaceMatchCase_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRange::ReplaceMatchCase)>::type;

struct ExRange_ReplaceMatchCase : public ExRange, public Aspose::Words::ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_ReplaceMatchCase_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExRange_ReplaceMatchCase, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ReplaceMatchCase(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_ReplaceMatchCase, ::testing::ValuesIn(ExRange_ReplaceMatchCase::TestCases()));

} // namespace gtest_test

void ExRange::ReplaceFindWholeWordsOnly(bool findWholeWordsOnly)
{
    //ExStart
    //ExFor:Range.Replace(String, String, FindReplaceOptions)
    //ExFor:FindReplaceOptions
    //ExFor:FindReplaceOptions.FindWholeWordsOnly
    //ExSummary:Shows how to toggle standalone word-only find-and-replace operations. 
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Jackson will meet you in Jacksonville.");
    
    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    
    // Set the "FindWholeWordsOnly" flag to "true" to replace the found text if it is not a part of another word.
    // Set the "FindWholeWordsOnly" flag to "false" to replace all text regardless of its surroundings.
    options->set_FindWholeWordsOnly(findWholeWordsOnly);
    
    doc->get_Range()->Replace(u"Jackson", u"Louis", options);
    
    ASSERT_EQ(findWholeWordsOnly ? System::String(u"Louis will meet you in Jacksonville.") : System::String(u"Louis will meet you in Louisville."), doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

using ExRange_ReplaceFindWholeWordsOnly_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRange::ReplaceFindWholeWordsOnly)>::type;

struct ExRange_ReplaceFindWholeWordsOnly : public ExRange, public Aspose::Words::ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_ReplaceFindWholeWordsOnly_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExRange_ReplaceFindWholeWordsOnly, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ReplaceFindWholeWordsOnly(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_ReplaceFindWholeWordsOnly, ::testing::ValuesIn(ExRange_ReplaceFindWholeWordsOnly::TestCases()));

} // namespace gtest_test

void ExRange::IgnoreDeleted(bool ignoreTextInsideDeleteRevisions)
{
    //ExStart
    //ExFor:FindReplaceOptions.IgnoreDeleted
    //ExSummary:Shows how to include or ignore text inside delete revisions during a find-and-replace operation.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Hello world!");
    builder->Writeln(u"Hello again!");
    
    // Start tracking revisions and remove the second paragraph, which will create a delete revision.
    // That paragraph will persist in the document until we accept the delete revision.
    doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
    doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->Remove();
    doc->StopTrackRevisions();
    
    ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_IsDeleteRevision());
    
    // We can use a "FindReplaceOptions" object to modify the find and replace process.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    
    // Set the "IgnoreDeleted" flag to "true" to get the find-and-replace
    // operation to ignore paragraphs that are delete revisions.
    // Set the "IgnoreDeleted" flag to "false" to get the find-and-replace
    // operation to also search for text inside delete revisions.
    options->set_IgnoreDeleted(ignoreTextInsideDeleteRevisions);
    
    doc->get_Range()->Replace(u"Hello", u"Greetings", options);
    
    ASSERT_EQ(ignoreTextInsideDeleteRevisions ? System::String(u"Greetings world!\rHello again!") : System::String(u"Greetings world!\rGreetings again!"), doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

using ExRange_IgnoreDeleted_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRange::IgnoreDeleted)>::type;

struct ExRange_IgnoreDeleted : public ExRange, public Aspose::Words::ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_IgnoreDeleted_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExRange_IgnoreDeleted, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreDeleted(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_IgnoreDeleted, ::testing::ValuesIn(ExRange_IgnoreDeleted::TestCases()));

} // namespace gtest_test

void ExRange::IgnoreInserted(bool ignoreTextInsideInsertRevisions)
{
    //ExStart
    //ExFor:FindReplaceOptions.IgnoreInserted
    //ExSummary:Shows how to include or ignore text inside insert revisions during a find-and-replace operation.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Hello world!");
    
    // Start tracking revisions and insert a paragraph. That paragraph will be an insert revision.
    doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
    builder->Writeln(u"Hello again!");
    doc->StopTrackRevisions();
    
    ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_IsInsertRevision());
    
    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    
    // Set the "IgnoreInserted" flag to "true" to get the find-and-replace
    // operation to ignore paragraphs that are insert revisions.
    // Set the "IgnoreInserted" flag to "false" to get the find-and-replace
    // operation to also search for text inside insert revisions.
    options->set_IgnoreInserted(ignoreTextInsideInsertRevisions);
    
    doc->get_Range()->Replace(u"Hello", u"Greetings", options);
    
    ASSERT_EQ(ignoreTextInsideInsertRevisions ? System::String(u"Greetings world!\rHello again!") : System::String(u"Greetings world!\rGreetings again!"), doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

using ExRange_IgnoreInserted_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRange::IgnoreInserted)>::type;

struct ExRange_IgnoreInserted : public ExRange, public Aspose::Words::ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_IgnoreInserted_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExRange_IgnoreInserted, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreInserted(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_IgnoreInserted, ::testing::ValuesIn(ExRange_IgnoreInserted::TestCases()));

} // namespace gtest_test

void ExRange::IgnoreFields(bool ignoreTextInsideFields)
{
    //ExStart
    //ExFor:FindReplaceOptions.IgnoreFields
    //ExSummary:Shows how to ignore text inside fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Hello world!");
    builder->InsertField(u"QUOTE", u"Hello again!");
    
    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    
    // Set the "IgnoreFields" flag to "true" to get the find-and-replace
    // operation to ignore text inside fields.
    // Set the "IgnoreFields" flag to "false" to get the find-and-replace
    // operation to also search for text inside fields.
    options->set_IgnoreFields(ignoreTextInsideFields);
    
    doc->get_Range()->Replace(u"Hello", u"Greetings", options);
    
    ASSERT_EQ(ignoreTextInsideFields ? System::String(u"Greetings world!\r\u0013QUOTE\u0014Hello again!\u0015") : System::String(u"Greetings world!\r\u0013QUOTE\u0014Greetings again!\u0015"), doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

using ExRange_IgnoreFields_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRange::IgnoreFields)>::type;

struct ExRange_IgnoreFields : public ExRange, public Aspose::Words::ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_IgnoreFields_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExRange_IgnoreFields, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreFields(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_IgnoreFields, ::testing::ValuesIn(ExRange_IgnoreFields::TestCases()));

} // namespace gtest_test

void ExRange::IgnoreFieldCodes(bool ignoreFieldCodes)
{
    //ExStart
    //ExFor:FindReplaceOptions.IgnoreFieldCodes
    //ExSummary:Shows how to ignore text inside field codes.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertField(u"INCLUDETEXT", u"Test IT!");
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    options->set_IgnoreFieldCodes(ignoreFieldCodes);
    
    // Replace 'T' in document ignoring text inside field code or not.
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"T"), u"*", options);
    std::cout << doc->GetText() << std::endl;
    
    ASSERT_EQ(ignoreFieldCodes ? System::String(u"\u0013INCLUDETEXT\u0014*est I*!\u0015") : System::String(u"\u0013INCLUDE*EX*\u0014*est I*!\u0015"), doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

using ExRange_IgnoreFieldCodes_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRange::IgnoreFieldCodes)>::type;

struct ExRange_IgnoreFieldCodes : public ExRange, public Aspose::Words::ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_IgnoreFieldCodes_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExRange_IgnoreFieldCodes, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreFieldCodes(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_IgnoreFieldCodes, ::testing::ValuesIn(ExRange_IgnoreFieldCodes::TestCases()));

} // namespace gtest_test

void ExRange::IgnoreFootnote(bool isIgnoreFootnotes)
{
    //ExStart
    //ExFor:FindReplaceOptions.IgnoreFootnotes
    //ExSummary:Shows how to ignore footnotes during a find-and-replace operation.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
    
    builder->InsertParagraph();
    
    builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Endnote, u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
    
    // Set the "IgnoreFootnotes" flag to "true" to get the find-and-replace
    // operation to ignore text inside footnotes.
    // Set the "IgnoreFootnotes" flag to "false" to get the find-and-replace
    // operation to also search for text inside footnotes.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    options->set_IgnoreFootnotes(isIgnoreFootnotes);
    doc->get_Range()->Replace(u"Lorem ipsum", u"Replaced Lorem ipsum", options);
    //ExEnd
    
    System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    
    for (auto&& para : System::IterateOver<Aspose::Words::Paragraph>(paragraphs))
    {
        ASSERT_EQ(u"Replaced Lorem ipsum", para->get_Runs()->idx_get(0)->get_Text());
    }
    
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Notes::Footnote>>> footnotes = doc->GetChildNodes(Aspose::Words::NodeType::Footnote, true)->LINQ_Cast<System::SharedPtr<Aspose::Words::Notes::Footnote> >()->LINQ_ToList();
    ASSERT_EQ(isIgnoreFootnotes ? System::String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.") : System::String(u"Replaced Lorem ipsum dolor sit amet, consectetur adipiscing elit."), footnotes->idx_get(0)->ToString(Aspose::Words::SaveFormat::Text).Trim());
    ASSERT_EQ(isIgnoreFootnotes ? System::String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.") : System::String(u"Replaced Lorem ipsum dolor sit amet, consectetur adipiscing elit."), footnotes->idx_get(1)->ToString(Aspose::Words::SaveFormat::Text).Trim());
}

namespace gtest_test
{

using ExRange_IgnoreFootnote_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRange::IgnoreFootnote)>::type;

struct ExRange_IgnoreFootnote : public ExRange, public Aspose::Words::ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_IgnoreFootnote_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExRange_IgnoreFootnote, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreFootnote(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_IgnoreFootnote, ::testing::ValuesIn(ExRange_IgnoreFootnote::TestCases()));

} // namespace gtest_test

void ExRange::IgnoreShapes()
{
    //ExStart
    //ExFor:FindReplaceOptions.IgnoreShapes
    //ExSummary:Shows how to ignore shapes while replacing text.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::Balloon, 200, 200);
    builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
    auto findReplaceOptions = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    findReplaceOptions->set_IgnoreShapes(true);
    builder->get_Document()->get_Range()->Replace(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.Lorem ipsum dolor sit amet, consectetur adipiscing elit.", u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.", findReplaceOptions);
    ASSERT_EQ(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit.", builder->get_Document()->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRange, IgnoreShapes)
{
    s_instance->IgnoreShapes();
}

} // namespace gtest_test

void ExRange::UpdateFieldsInRange()
{
    //ExStart
    //ExFor:Range.UpdateFields
    //ExSummary:Shows how to update all the fields in a range.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertField(u" DOCPROPERTY Category");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakEvenPage);
    builder->InsertField(u" DOCPROPERTY Category");
    
    // The above DOCPROPERTY fields will display the value of this built-in document property.
    doc->get_BuiltInDocumentProperties()->set_Category(u"MyCategory");
    
    // If we update the value of a document property, we will need to update all the DOCPROPERTY fields to display it.
    ASSERT_EQ(System::String::Empty, doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
    ASSERT_EQ(System::String::Empty, doc->get_Range()->get_Fields()->idx_get(1)->get_Result());
    
    // Update all the fields that are in the range of the first section.
    doc->get_FirstSection()->get_Range()->UpdateFields();
    
    ASSERT_EQ(u"MyCategory", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
    ASSERT_EQ(System::String::Empty, doc->get_Range()->get_Fields()->idx_get(1)->get_Result());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRange, UpdateFieldsInRange)
{
    s_instance->UpdateFieldsInRange();
}

} // namespace gtest_test

void ExRange::ReplaceWithString()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"This one is sad.");
    builder->Writeln(u"That one is mad.");
    
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    options->set_MatchCase(false);
    options->set_FindWholeWordsOnly(true);
    
    doc->get_Range()->Replace(u"sad", u"bad", options);
    
    doc->Save(get_ArtifactsDir() + u"Range.ReplaceWithString.docx");
}

namespace gtest_test
{

TEST_F(ExRange, ReplaceWithString)
{
    s_instance->ReplaceWithString();
}

} // namespace gtest_test

void ExRange::ReplaceWithRegex()
{
    //ExStart
    //ExFor:Range.Replace(Regex, String)
    //ExSummary:Shows how to replace all occurrences of a regular expression pattern with other text.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"I decided to get the curtains in gray, ideal for the grey-accented room.");
    
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"gr(a|e)y"), u"lavender");
    
    ASSERT_EQ(u"I decided to get the curtains in lavender, ideal for the lavender-accented room.", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRange, ReplaceWithRegex)
{
    s_instance->ReplaceWithRegex();
}

} // namespace gtest_test

void ExRange::ReplaceWithCallback()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(System::String(u"Our new location in New York City is opening tomorrow. ") + u"Hope to see all our NYC-based customers at the opening!");
    
    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    
    // Set a callback that tracks any replacements that the "Replace" method will make.
    auto logger = System::MakeObject<Aspose::Words::ApiExamples::ExRange::TextFindAndReplacementLogger>();
    options->set_ReplacingCallback(logger);
    
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"New York City|NYC"), u"Washington", options);
    
    ASSERT_EQ(System::String(u"Our new location in (Old value:\"New York City\") Washington is opening tomorrow. ") + u"Hope to see all our (Old value:\"NYC\") Washington-based customers at the opening!", doc->GetText().Trim());
    
    ASSERT_EQ(System::String(u"\"New York City\" converted to \"Washington\" 20 characters into a Run node.\r\n") + u"\"NYC\" converted to \"Washington\" 42 characters into a Run node.", logger->GetLog().Trim());
}

namespace gtest_test
{

TEST_F(ExRange, ReplaceWithCallback)
{
    s_instance->ReplaceWithCallback();
}

} // namespace gtest_test

void ExRange::ConvertNumbersToHexadecimal()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Arial");
    builder->Writeln(System::String(u"Numbers that the find-and-replace operation will convert to hexadecimal and highlight:\n") + u"123, 456, 789 and 17379.");
    
    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    
    // Set the "HighlightColor" property to a background color that we want to apply to the operation's resulting text.
    options->get_ApplyFont()->set_HighlightColor(System::Drawing::Color::get_LightGray());
    
    auto numberHexer = System::MakeObject<Aspose::Words::ApiExamples::ExRange::NumberHexer>();
    options->set_ReplacingCallback(numberHexer);
    
    int32_t replacementCount = doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"[0-9]+"), u"", options);
    
    std::cout << numberHexer->GetLog() << std::endl;
    
    ASSERT_EQ(4, replacementCount);
    ASSERT_EQ(System::String(u"Numbers that the find-and-replace operation will convert to hexadecimal and highlight:\r") + u"0x7B, 0x1C8, 0x315 and 0x43E3.", doc->GetText().Trim());
    ASSERT_EQ(4, doc->GetChildNodes(Aspose::Words::NodeType::Run, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Run> >()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Run>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Run> r)>>([](System::SharedPtr<Aspose::Words::Run> r) -> bool
    {
        return r->get_Font()->get_HighlightColor().ToArgb() == System::Drawing::Color::get_LightGray().ToArgb();
    }))));
}

namespace gtest_test
{

TEST_F(ExRange, ConvertNumbersToHexadecimal)
{
    s_instance->ConvertNumbersToHexadecimal();
}

} // namespace gtest_test

void ExRange::ApplyParagraphFormat()
{
    //ExStart
    //ExFor:FindReplaceOptions.ApplyParagraphFormat
    //ExFor:Range.Replace(String, String)
    //ExSummary:Shows how to add formatting to paragraphs in which a find-and-replace operation has found matches.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Every paragraph that ends with a full stop like this one will be right aligned.");
    builder->Writeln(u"This one will not!");
    builder->Write(u"This one also will.");
    
    System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Left, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Alignment());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Left, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Alignment());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Left, paragraphs->idx_get(2)->get_ParagraphFormat()->get_Alignment());
    
    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    
    // Set the "Alignment" property to "ParagraphAlignment.Right" to right-align every paragraph
    // that contains a match that the find-and-replace operation finds.
    options->get_ApplyParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Right);
    
    // Replace every full stop that is right before a paragraph break with an exclamation point.
    int32_t count = doc->get_Range()->Replace(u".&p", u"!&p", options);
    
    ASSERT_EQ(2, count);
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Right, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Alignment());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Left, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Alignment());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Right, paragraphs->idx_get(2)->get_ParagraphFormat()->get_Alignment());
    ASSERT_EQ(System::String(u"Every paragraph that ends with a full stop like this one will be right aligned!\r") + u"This one will not!\r" + u"This one also will!", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRange, ApplyParagraphFormat)
{
    s_instance->ApplyParagraphFormat();
}

} // namespace gtest_test

void ExRange::DeleteSelection()
{
    //ExStart
    //ExFor:Node.Range
    //ExFor:Range.Delete
    //ExSummary:Shows how to delete all the nodes from a range.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Add text to the first section in the document, and then add another section.
    builder->Write(u"Section 1. ");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakContinuous);
    builder->Write(u"Section 2.");
    
    ASSERT_EQ(u"Section 1. \fSection 2.", doc->GetText().Trim());
    
    // Remove the first section entirely by removing all the nodes
    // within its range, including the section itself.
    doc->get_Sections()->idx_get(0)->get_Range()->Delete();
    
    ASSERT_EQ(1, doc->get_Sections()->get_Count());
    ASSERT_EQ(u"Section 2.", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRange, DeleteSelection)
{
    s_instance->DeleteSelection();
}

} // namespace gtest_test

void ExRange::RangesGetText()
{
    //ExStart
    //ExFor:Range
    //ExFor:Range.Text
    //ExSummary:Shows how to get the text contents of all the nodes that a range covers.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Hello world!");
    
    ASSERT_EQ(u"Hello world!", doc->get_Range()->get_Text().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExRange, RangesGetText)
{
    s_instance->RangesGetText();
}

} // namespace gtest_test

void ExRange::UseLegacyOrder(bool useLegacyOrder)
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert three runs which we can search for using a regex pattern.
    // Place one of those runs inside a text box.
    builder->Writeln(u"[tag 1]");
    System::SharedPtr<Aspose::Words::Drawing::Shape> textBox = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 50);
    builder->Writeln(u"[tag 2]");
    builder->MoveTo(textBox->get_FirstParagraph());
    builder->Write(u"[tag 3]");
    
    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    
    // Assign a custom callback to the "ReplacingCallback" property.
    auto callback = System::MakeObject<Aspose::Words::ApiExamples::ExRange::TextReplacementTracker>();
    options->set_ReplacingCallback(callback);
    
    // If we set the "UseLegacyOrder" property to "true", the
    // find-and-replace operation will go through all the runs outside of a text box
    // before going through the ones inside a text box.
    // If we set the "UseLegacyOrder" property to "false", the
    // find-and-replace operation will go over all the runs in a range in sequential order.
    options->set_UseLegacyOrder(useLegacyOrder);
    
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"\\[tag \\d*\\]"), u"", options);
    
    System::SharedPtr<System::Collections::Generic::List<System::String>> expected;
    if (useLegacyOrder)
    {
        expected = [&]{ System::String init_0[] = {u"[tag 1]", u"[tag 3]", u"[tag 2]"}; auto list_0 = System::MakeObject<System::Collections::Generic::List<System::String>>(); list_0->AddInitializer(3, init_0); return list_0; }();
    }
    else
    {
        expected = [&]{ System::String init_1[] = {u"[tag 1]", u"[tag 2]", u"[tag 3]"}; auto list_1 = System::MakeObject<System::Collections::Generic::List<System::String>>(); list_1->AddInitializer(3, init_1); return list_1; }();
    }
    ASPOSE_ASSERT_EQ(expected, callback->get_Matches());
    
}

namespace gtest_test
{

using ExRange_UseLegacyOrder_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRange::UseLegacyOrder)>::type;

struct ExRange_UseLegacyOrder : public ExRange, public Aspose::Words::ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_UseLegacyOrder_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExRange_UseLegacyOrder, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UseLegacyOrder(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_UseLegacyOrder, ::testing::ValuesIn(ExRange_UseLegacyOrder::TestCases()));

} // namespace gtest_test

void ExRange::UseSubstitutions(bool useSubstitutions)
{
    //ExStart
    //ExFor:FindReplaceOptions.UseSubstitutions
    //ExSummary:Shows how to replace the text with substitutions.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"John sold a car to Paul.");
    builder->Writeln(u"Jane sold a house to Joe.");
    
    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    
    // Set the "UseSubstitutions" property to "true" to get
    // the find-and-replace operation to recognize substitution elements.
    // Set the "UseSubstitutions" property to "false" to ignore substitution elements.
    options->set_UseSubstitutions(useSubstitutions);
    
    auto regex = System::MakeObject<System::Text::RegularExpressions::Regex>(u"([A-z]+) sold a ([A-z]+) to ([A-z]+)");
    doc->get_Range()->Replace(regex, u"$3 bought a $2 from $1", options);
    
    ASSERT_EQ(useSubstitutions ? System::String(u"Paul bought a car from John.\rJoe bought a house from Jane.") : System::String(u"$3 bought a $2 from $1.\r$3 bought a $2 from $1."), doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

using ExRange_UseSubstitutions_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRange::UseSubstitutions)>::type;

struct ExRange_UseSubstitutions : public ExRange, public Aspose::Words::ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_UseSubstitutions_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExRange_UseSubstitutions, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UseSubstitutions(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_UseSubstitutions, ::testing::ValuesIn(ExRange_UseSubstitutions::TestCases()));

} // namespace gtest_test

void ExRange::InsertDocumentAtReplace()
{
    auto mainDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document insertion destination.docx");
    
    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    options->set_ReplacingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExRange::InsertDocumentAtReplaceHandler>());
    
    mainDoc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"\\[MY_DOCUMENT\\]"), u"", options);
    mainDoc->Save(get_ArtifactsDir() + u"InsertDocument.InsertDocumentAtReplace.docx");
    
    TestInsertDocumentAtReplace(System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"InsertDocument.InsertDocumentAtReplace.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExRange, InsertDocumentAtReplace)
{
    s_instance->InsertDocumentAtReplace();
}

} // namespace gtest_test

void ExRange::Direction(Aspose::Words::Replacing::FindReplaceDirection findReplaceDirection)
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert three runs which we can search for using a regex pattern.
    // Place one of those runs inside a text box.
    builder->Writeln(u"Match 1.");
    builder->Writeln(u"Match 2.");
    builder->Writeln(u"Match 3.");
    builder->Writeln(u"Match 4.");
    
    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    
    // Assign a custom callback to the "ReplacingCallback" property.
    auto callback = System::MakeObject<Aspose::Words::ApiExamples::ExRange::TextReplacementRecorder>();
    options->set_ReplacingCallback(callback);
    
    // Set the "Direction" property to "FindReplaceDirection.Backward" to get the find-and-replace
    // operation to start from the end of the range, and traverse back to the beginning.
    // Set the "Direction" property to "FindReplaceDirection.Forward" to get the find-and-replace
    // operation to start from the beginning of the range, and traverse to the end.
    options->set_Direction(findReplaceDirection);
    
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"Match \\d*"), u"Replacement", options);
    
    ASSERT_EQ(System::String(u"Replacement.\r") + u"Replacement.\r" + u"Replacement.\r" + u"Replacement.", doc->GetText().Trim());
    
    switch (findReplaceDirection)
    {
        case Aspose::Words::Replacing::FindReplaceDirection::Forward:
            ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"Match 1", u"Match 2", u"Match 3", u"Match 4"}), callback->get_Matches());
            break;
        
        case Aspose::Words::Replacing::FindReplaceDirection::Backward:
            ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"Match 4", u"Match 3", u"Match 2", u"Match 1"}), callback->get_Matches());
            break;
        
    }
}

namespace gtest_test
{

using ExRange_Direction_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRange::Direction)>::type;

struct ExRange_Direction : public ExRange, public Aspose::Words::ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_Direction_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Replacing::FindReplaceDirection::Backward),
            std::make_tuple(Aspose::Words::Replacing::FindReplaceDirection::Forward),
        };
    }
};

TEST_P(ExRange_Direction, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Direction(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_Direction, ::testing::ValuesIn(ExRange_Direction::TestCases()));

} // namespace gtest_test

void ExRange::MatchEndNode()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"1");
    builder->Writeln(u"2");
    builder->Writeln(u"3");
    
    auto replacingCallback = System::MakeObject<Aspose::Words::ApiExamples::ExRange::ReplacingCallback>();
    auto opts = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    opts->set_ReplacingCallback(replacingCallback);
    
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"1[\\s\\S]*3"), u"X", opts);
    ASSERT_EQ(u"1", replacingCallback->get_StartNodeText());
    ASSERT_EQ(u"3", replacingCallback->get_EndNodeText());
}

namespace gtest_test
{

TEST_F(ExRange, MatchEndNode)
{
    s_instance->MatchEndNode();
}

} // namespace gtest_test

void ExRange::IgnoreOfficeMath(bool isIgnoreOfficeMath)
{
    //ExStart:IgnoreOfficeMath
    //GistId:571cc6e23284a2ec075d15d4c32e3bbf
    //ExFor:FindReplaceOptions.IgnoreOfficeMath
    //ExSummary:Shows how to find and replace text within OfficeMath.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    ASSERT_EQ(u"i+b-c≥iM+bM-cM", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText().Trim());
    
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    options->set_IgnoreOfficeMath(isIgnoreOfficeMath);
    doc->get_Range()->Replace(u"b", u"x", options);
    
    if (isIgnoreOfficeMath)
    {
        ASSERT_EQ(u"i+b-c≥iM+bM-cM", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText().Trim());
    }
    else
    {
        ASSERT_EQ(u"i+x-c≥iM+xM-cM", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText().Trim());
    }
    //ExEnd:IgnoreOfficeMath
}

namespace gtest_test
{

using ExRange_IgnoreOfficeMath_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRange::IgnoreOfficeMath)>::type;

struct ExRange_IgnoreOfficeMath : public ExRange, public Aspose::Words::ApiExamples::ExRange, public ::testing::WithParamInterface<ExRange_IgnoreOfficeMath_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExRange_IgnoreOfficeMath, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreOfficeMath(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRange_IgnoreOfficeMath, ::testing::ValuesIn(ExRange_IgnoreOfficeMath::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
