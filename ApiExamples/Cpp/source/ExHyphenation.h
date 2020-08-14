#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/collections/dictionary.h>
#include <Aspose.Words.Cpp/Layout/Hyphenation/IHyphenationCallback.h>

#include "ApiExampleBase.h"

namespace ApiExamples {

class ExHyphenation : public ApiExampleBase
{
private:

    /// <summary>
    /// Associates ISO language codes with custom local system dictionary files for their respective languages
    /// </summary>
    class CustomHyphenationDictionaryRegister : public Aspose::Words::IHyphenationCallback
    {
        typedef CustomHyphenationDictionaryRegister ThisType;
        typedef Aspose::Words::IHyphenationCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        CustomHyphenationDictionaryRegister();
        
        void RequestDictionary(System::String language) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Collections::Generic::Dictionary<System::String, System::String>> mHyphenationDictionaryFiles;
        
    };
    
    
public:

    void Dictionary();
    void RegisterDictionary();
    
};

} // namespace ApiExamples


