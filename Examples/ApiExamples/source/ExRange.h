#pragma once

#include <system/text/string_builder.h>
#include <system/string.h>
#include <system/collections/list.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceDirection.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Replacing;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExRange : public ApiExampleBase
{
    typedef ExRange ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Maintains a log of every text replacement done by a find-and-replace operation
    /// and notes the original matched text's value.
    /// </summary>
    class TextFindAndReplacementLogger : public IReplacingCallback
    {
        typedef TextFindAndReplacementLogger ThisType;
        typedef IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String GetLog();
        
        TextFindAndReplacementLogger();
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mLog;
        
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override;
        
    };
    
    /// <summary>
    /// Replaces numeric find-and-replacement matches with their hexadecimal equivalents.
    /// Maintains a log of every replacement.
    /// </summary>
    class NumberHexer : public IReplacingCallback
    {
        typedef NumberHexer ThisType;
        typedef IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override;
        System::String GetLog();
        
        NumberHexer();
        
    private:
    
        int32_t mCurrentReplacementNumber;
        System::SharedPtr<System::Text::StringBuilder> mLog;
        
    };
    
    /// <summary>
    /// Records the order of all matches that occur during a find-and-replace operation.
    /// </summary>
    class TextReplacementTracker : public IReplacingCallback
    {
        typedef TextReplacementTracker ThisType;
        typedef IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<System::Collections::Generic::List<System::String>> get_Matches() const;
        
        TextReplacementTracker();
        
    private:
    
        System::SharedPtr<System::Collections::Generic::List<System::String>> mMatches;
        
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> e) override;
        
    };
    
    class InsertDocumentAtReplaceHandler : public IReplacingCallback
    {
        typedef InsertDocumentAtReplaceHandler ThisType;
        typedef IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    private:
    
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override;
        
    };
    
    /// <summary>
    /// Records all matches that occur during a find-and-replace operation in the order that they take place.
    /// </summary>
    class TextReplacementRecorder : public IReplacingCallback
    {
        typedef TextReplacementRecorder ThisType;
        typedef IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<System::Collections::Generic::List<System::String>> get_Matches() const;
        
        TextReplacementRecorder();
        
    private:
    
        System::SharedPtr<System::Collections::Generic::List<System::String>> mMatches;
        
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> e) override;
        
    };
    
    /// <summary>
    /// The replacing callback.
    /// </summary>
    class ReplacingCallback : public IReplacingCallback
    {
        typedef ReplacingCallback ThisType;
        typedef IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String get_StartNodeText() const;
        System::String get_EndNodeText() const;
        
    private:
    
        System::String pr_StartNodeText;
        
        void set_StartNodeText(System::String value);
        
        System::String pr_EndNodeText;
        
        void set_EndNodeText(System::String value);
        
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> e) override;
        
    };
    
    
public:

    void Replace();
    void ReplaceMatchCase(bool matchCase);
    void ReplaceFindWholeWordsOnly(bool findWholeWordsOnly);
    void IgnoreDeleted(bool ignoreTextInsideDeleteRevisions);
    void IgnoreInserted(bool ignoreTextInsideInsertRevisions);
    void IgnoreFields(bool ignoreTextInsideFields);
    void IgnoreFieldCodes(bool ignoreFieldCodes);
    void IgnoreFootnote(bool isIgnoreFootnotes);
    void IgnoreShapes();
    void UpdateFieldsInRange();
    void ReplaceWithString();
    void ReplaceWithRegex();
    //ExStart
    //ExFor:FindReplaceOptions.ReplacingCallback
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:ReplacingArgs.Replacement
    //ExFor:IReplacingCallback
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExSummary:Shows how to replace all occurrences of a regular expression pattern with another string, while tracking all such replacements.
    void ReplaceWithCallback();
    //ExEnd
    //ExStart
    //ExFor:FindReplaceOptions.ApplyFont
    //ExFor:FindReplaceOptions.ReplacingCallback
    //ExFor:ReplacingArgs.GroupIndex
    //ExFor:ReplacingArgs.GroupName
    //ExFor:ReplacingArgs.Match
    //ExFor:ReplacingArgs.MatchOffset
    //ExSummary:Shows how to apply a different font to new content via FindReplaceOptions.
    void ConvertNumbersToHexadecimal();
    //ExEnd
    void ApplyParagraphFormat();
    void DeleteSelection();
    void RangesGetText();
    //ExStart
    //ExFor:FindReplaceOptions.UseLegacyOrder
    //ExSummary:Shows how to change the searching order of nodes when performing a find-and-replace text operation.
    void UseLegacyOrder(bool useLegacyOrder);
    //ExEnd
    void UseSubstitutions(bool useSubstitutions);
    //ExStart
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:IReplacingCallback
    //ExFor:ReplaceAction
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExFor:ReplacingArgs.MatchNode
    //ExSummary:Shows how to insert an entire document's contents as a replacement of a match in a find-and-replace operation.
    void InsertDocumentAtReplace();
    //ExStart
    //ExFor:FindReplaceOptions.Direction
    //ExFor:FindReplaceDirection
    //ExSummary:Shows how to determine which direction a find-and-replace operation traverses the document in.
    void Direction(Aspose::Words::Replacing::FindReplaceDirection findReplaceDirection);
    //ExEnd
    //ExStart:MatchEndNode
    //GistId:67c1d01ce69d189983b497fd497a7768
    //ExFor:ReplacingArgs.MatchEndNode
    //ExSummary:Shows how to get match end node.
    void MatchEndNode();
    //ExEnd:MatchEndNode
    void IgnoreOfficeMath(bool isIgnoreOfficeMath);
    
protected:

    /// <summary>
    /// Inserts all the nodes of another document after a paragraph or table.
    /// </summary>
    static void InsertDocument(System::SharedPtr<Aspose::Words::Node> insertionDestination, System::SharedPtr<Aspose::Words::Document> docToInsert);
    //ExEnd
    static void TestInsertDocumentAtReplace(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


