#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Replacing;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExHeaderFooter : public ApiExampleBase
{
    typedef ExHeaderFooter ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// During a find-and-replace operation, records the contents of every node that has text that the operation 'finds',
    /// in the state it is in before the replacement takes place.
    /// This will display the order in which the text replacement operation traverses nodes.
    /// </summary>
    class ReplaceLog : public IReplacingCallback
    {
        typedef ReplaceLog ThisType;
        typedef IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String get_Text();
        
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override;
        
        ReplaceLog();
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mTextBuilder;
        
    };
    
    
public:

    void Create();
    void Link();
    void RemoveFooters();
    void ExportMode();
    void ReplaceText();
    void Order(bool differentFirstPageHeaderFooter);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


