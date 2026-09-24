#pragma once

#include <system/text/string_builder.h>
#include <system/string.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/AbsolutePositionTab.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExAbsolutePositionTab : public ApiExampleBase
{
    typedef ExAbsolutePositionTab ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Collects the text contents of all runs in the visited document. Replaces all absolute tab characters with ordinary tabs.
    /// </summary>
    class DocTextExtractor : public DocumentVisitor
    {
        typedef DocTextExtractor ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        DocTextExtractor();
        
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when an AbsolutePositionTab node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitAbsolutePositionTab(System::SharedPtr<Aspose::Words::AbsolutePositionTab> tab) override;
        /// <summary>
        /// Plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Adds text to the current output. Honors the enabled/disabled output flag.
        /// </summary>
        void AppendText(System::String text);
        
    };
    
    
public:

    //ExStart
    //ExFor:AbsolutePositionTab
    //ExFor:AbsolutePositionTab.Accept(DocumentVisitor)
    //ExFor:DocumentVisitor.VisitAbsolutePositionTab
    //ExFor:Body.Accept(DocumentVisitor)
    //ExFor:Body.AcceptStart(DocumentVisitor)
    //ExFor:Body.AcceptEnd(DocumentVisitor)
    //ExFor:VisitorAction
    //ExSummary:Shows how to process absolute position tab characters with a document visitor.
    void DocumentToTxt();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


