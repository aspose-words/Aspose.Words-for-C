#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Text/AbsolutePositionTab.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <system/exceptions.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/string_builder.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;

namespace ApiExamples {

class ExAbsolutePositionTab : public ApiExampleBase
{
public:
    //ExStart
    //ExFor:AbsolutePositionTab
    //ExFor:AbsolutePositionTab.Accept(DocumentVisitor)
    //ExFor:DocumentVisitor.VisitAbsolutePositionTab
    //ExSummary:Shows how to process absolute position tab characters with a document visitor.
    void DocumentToTxt()
    {
        auto doc = MakeObject<Document>(MyDir + u"Absolute position tab.docx");

        // Extract the text contents of our document by accepting this custom document visitor.
        auto myDocTextExtractor = MakeObject<ExAbsolutePositionTab::DocTextExtractor>();
        doc->get_FirstSection()->get_Body()->Accept(myDocTextExtractor);

        // The absolute position tab, which has no equivalent in string form, has been explicitly converted to a tab character.
        ASSERT_EQ(u"Before AbsolutePositionTab\tAfter AbsolutePositionTab", myDocTextExtractor->GetText());

        // An AbsolutePositionTab can accept a DocumentVisitor by itself too.
        auto absPositionTab =
            System::DynamicCast<AbsolutePositionTab>(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetChild(NodeType::SpecialChar, 0, true));

        myDocTextExtractor = MakeObject<ExAbsolutePositionTab::DocTextExtractor>();
        absPositionTab->Accept(myDocTextExtractor);

        ASSERT_EQ(u"\t", myDocTextExtractor->GetText());
    }

    /// <summary>
    /// Collects the text contents of all runs in the visited document, and represents all absolute tab characters as ordinary tabs.
    /// </summary>
    class DocTextExtractor : public DocumentVisitor
    {
    public:
        DocTextExtractor()
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        VisitorAction VisitRun(SharedPtr<Run> run) override
        {
            AppendText(run->get_Text());
            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when an AbsolutePositionTab node is encountered in the document.
        /// </summary>
        VisitorAction VisitAbsolutePositionTab(SharedPtr<AbsolutePositionTab> tab) override
        {
            mBuilder->Append(u"\t");
            return VisitorAction::Continue;
        }

        /// <summary>
        /// Plain text of the document that was accumulated by the visitor.
        /// </summary>
        String GetText()
        {
            return mBuilder->ToString();
        }

    private:
        SharedPtr<System::Text::StringBuilder> mBuilder;

        /// <summary>
        /// Adds text to the current output. Honors the enabled/disabled output flag.
        /// </summary>
        void AppendText(String text)
        {
            mBuilder->Append(text);
        }
    };
    //ExEnd
};

} // namespace ApiExamples
