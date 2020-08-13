#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Replacing { enum class ReplaceAction; } } }
namespace Aspose { namespace Words { namespace Replacing { class ReplacingArgs; } } }
namespace Aspose { namespace Words { class Section; } }

namespace ApiExamples {

class ExHeaderFooter : public ApiExampleBase
{
private:

    class ReplaceLog : public Aspose::Words::Replacing::IReplacingCallback
    {
        typedef ReplaceLog ThisType;
        typedef Aspose::Words::Replacing::IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String get_Text();
        
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override;
        void ClearText();
        
        ReplaceLog();
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mTextBuilder;
        
    };
    
    
public:

    void HeaderFooterCreate();
    void HeaderFooterLink();
    void RemoveFooters();
    void SetExportHeadersFootersMode();
    void ReplaceText();
    void HeaderFooterOrder();
    void Primer();
    
protected:

    /// <summary>
    /// Clones and copies headers/footers form the previous section to the specified section.
    /// </summary>
    static void CopyHeadersFootersFromPreviousSection(System::SharedPtr<Aspose::Words::Section> section);
    
};

} // namespace ApiExamples


