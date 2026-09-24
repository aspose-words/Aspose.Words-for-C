#pragma once

#include <system/text/string_builder.h>
#include <system/string.h>
#include <gtest/gtest.h>
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
    //ExStart
    //ExFor:IReplacingCallback
    //ExFor:PageSetup.DifferentFirstPageHeaderFooter
    //ExFor:FindReplaceOptions.#ctor(IReplacingCallback)
    //ExSummary:Shows how to track the order in which a text replacement operation traverses nodes.
    void Order(bool differentFirstPageHeaderFooter);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


