#pragma once

#include <system/string.h>
#include <system/collections/dictionary.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Layout/Hyphenation/IHyphenationCallback.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExHyphenation : public ApiExampleBase
{
    typedef ExHyphenation ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Associates ISO language codes with local system filenames for hyphenation dictionary files.
    /// </summary>
    class CustomHyphenationDictionaryRegister : public IHyphenationCallback
    {
        typedef CustomHyphenationDictionaryRegister ThisType;
        typedef IHyphenationCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        CustomHyphenationDictionaryRegister();
        
        void RequestDictionary(System::String language) override;
        
    private:
    
        System::SharedPtr<System::Collections::Generic::Dictionary<System::String, System::String>> mHyphenationDictionaryFiles;
        
    };
    
    
public:

    void Dictionary();
    //ExStart
    //ExFor:Hyphenation
    //ExFor:Hyphenation.Callback
    //ExFor:Hyphenation.RegisterDictionary(String, Stream)
    //ExFor:Hyphenation.RegisterDictionary(String, String)
    //ExFor:Hyphenation.WarningCallback
    //ExFor:IHyphenationCallback
    //ExFor:IHyphenationCallback.RequestDictionary(String)
    //ExSummary:Shows how to open and register a dictionary from a file.
    void RegisterDictionary();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


