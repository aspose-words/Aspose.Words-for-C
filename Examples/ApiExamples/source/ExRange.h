#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <system/collections/list.h>
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
    void ReplaceWithCallback();
    void ConvertNumbersToHexadecimal();
    void ApplyParagraphFormat();
    void DeleteSelection();
    void RangesGetText();
    void UseLegacyOrder(bool useLegacyOrder);
    void UseSubstitutions(bool useSubstitutions);
    void InsertDocumentAtReplace();
    void Direction(Aspose::Words::Replacing::FindReplaceDirection findReplaceDirection);
    void MatchEndNode();
    void IgnoreOfficeMath(bool isIgnoreOfficeMath);
    
protected:

    /// <summary>
    /// Inserts all the nodes of another document after a paragraph or table.
    /// </summary>
    static void InsertDocument(System::SharedPtr<Aspose::Words::Node> insertionDestination, System::SharedPtr<Aspose::Words::Document> docToInsert);
    static void TestInsertDocumentAtReplace(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


