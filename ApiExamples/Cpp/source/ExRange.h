#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/test_tools/method_argument_tuple.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Replacing { enum class ReplaceAction; } } }
namespace Aspose { namespace Words { namespace Replacing { class ReplacingArgs; } } }
namespace Aspose { namespace Words { class Node; } }
namespace Aspose { namespace Words { class Document; } }

namespace ApiExamples {

class ExRange : public ApiExampleBase
{
private:

    class ReplaceWithHtmlEvaluator : public Aspose::Words::Replacing::IReplacingCallback
    {
        typedef ReplaceWithHtmlEvaluator ThisType;
        typedef Aspose::Words::Replacing::IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override;
        
    };
    
    /// <summary>
    /// Replaces arabic numbers with hexadecimal equivalents and appends the number of each replacement.
    /// </summary>
    class NumberHexer : public Aspose::Words::Replacing::IReplacingCallback
    {
        typedef NumberHexer ThisType;
        typedef Aspose::Words::Replacing::IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override;
        
        NumberHexer();
        
    private:
    
        int32_t mCurrentReplacementNumber;
        
    };
    
    class InsertDocumentAtReplaceHandler : public Aspose::Words::Replacing::IReplacingCallback
    {
        typedef InsertDocumentAtReplaceHandler ThisType;
        typedef Aspose::Words::Replacing::IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override;
        
    };
    
    
public:

    void ReplaceSimple();
    void IgnoreDeleted(bool isIgnoreDeleted);
    void IgnoreInserted(bool isIgnoreInserted);
    void IgnoreFields(bool isIgnoreFields);
    void UpdateFieldsInRange();
    void ReplaceWithString();
    void ReplaceWithRegex();
    void ReplaceWithInsertHtml();
    void ReplaceNumbersAsHex();
    void ApplyParagraphFormat();
    void DeleteSelection();
    void RangesGetText();
    void InsertDocumentAtReplace();
    
protected:

    void TestInsertDocumentAtReplace(System::SharedPtr<Aspose::Words::Document> doc);
    
private:

    /// <summary>
    /// Inserts content of the external document after the specified node.
    /// </summary>
    static void InsertDocument(System::SharedPtr<Aspose::Words::Node> insertionDestination, System::SharedPtr<Aspose::Words::Document> docToInsert);
    
};

} // namespace ApiExamples


