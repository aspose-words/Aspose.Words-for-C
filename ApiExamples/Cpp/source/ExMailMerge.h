#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: System.Data is not supported

#include <system/test_tools/method_argument_tuple.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/MailMerge/IMailMergeCallback.h>

#include "ApiExampleBase.h"

namespace ApiExamples {

class ExMailMerge : public ApiExampleBase
{
private:

    class MailMergeCallbackStub : public Aspose::Words::MailMerging::IMailMergeCallback
    {
        typedef MailMergeCallbackStub ThisType;
        typedef Aspose::Words::MailMerging::IMailMergeCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        int32_t get_TagsReplacedCounter();
        
        void TagsReplaced() override;
        
        MailMergeCallbackStub();
        
    private:
    
        int32_t pr_TagsReplacedCounter;
        
        void set_TagsReplacedCounter(int32_t value);
        
    };
    
    
public:

    void MailMergeRegionInfo();
    void TrimWhiteSpaces(bool doTrimWhitespaces);
    void MailMergeGetFieldNames();
    void DeleteFields();
    void RemoveContainingFields();
    void RemoveUnusedFields();
    void RemoveEmptyParagraphs();
    void RemoveColonBetweenEmptyMergeFields(System::String punctuationMark, bool isCleanupParagraphsWithPunctuationMarks, System::String resultText);
    void GetFieldNames();
    void UseNonMergeFields();
    void TestMailMergeGetRegionsHierarchy();
    void TestTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption();
    void GetRegionsByName();
    
};

} // namespace ApiExamples


