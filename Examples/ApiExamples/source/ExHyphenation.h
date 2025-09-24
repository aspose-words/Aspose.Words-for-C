#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/collections/dictionary.h>
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
    void RegisterDictionary();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


