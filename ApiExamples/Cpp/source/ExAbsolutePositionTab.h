#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { enum class VisitorAction; } }
namespace Aspose { namespace Words { class Run; } }
namespace Aspose { namespace Words { class AbsolutePositionTab; } }

namespace ApiExamples {

class ExAbsolutePositionTab : public ApiExampleBase
{
public:

    /// <summary>
    /// Visitor implementation that simply collects the Runs and AbsolutePositionTabs of a document as plain text. 
    /// </summary>
    class DocToTxtWriter : public Aspose::Words::DocumentVisitor
    {
        typedef DocToTxtWriter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        DocToTxtWriter();
        
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when an AbsolutePositionTab node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitAbsolutePositionTab(System::SharedPtr<Aspose::Words::AbsolutePositionTab> tab) override;
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Adds text to the current output. Honors the enabled/disabled output flag.
        /// </summary>
        void AppendText(System::String text);
        
    };
    
    
public:

    void DocumentToTxt();
    
};

} // namespace ApiExamples


