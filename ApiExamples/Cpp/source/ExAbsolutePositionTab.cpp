// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExAbsolutePositionTab.h"

#include <testing/test_predicates.h>
#include <system/text/string_builder.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/AbsolutePositionTab.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3952568751u, ::ApiExamples::ExAbsolutePositionTab::DocToTxtWriter, ThisTypeBaseTypesInfo);

ExAbsolutePositionTab::DocToTxtWriter::DocToTxtWriter()
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
}

Aspose::Words::VisitorAction ExAbsolutePositionTab::DocToTxtWriter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    AppendText(run->get_Text());
    // Let the visitor continue visiting other nodes.
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExAbsolutePositionTab::DocToTxtWriter::VisitAbsolutePositionTab(SharedPtr<Aspose::Words::AbsolutePositionTab> tab)
{
    // We'll treat the AbsolutePositionTab as a regular tab in this case
    mBuilder->Append(u"\t");
    return Aspose::Words::VisitorAction::Continue;
}

void ExAbsolutePositionTab::DocToTxtWriter::AppendText(String text)
{
    mBuilder->Append(text);
}

String ExAbsolutePositionTab::DocToTxtWriter::GetText()
{
    return mBuilder->ToString();
}

System::Object::shared_members_type ApiExamples::ExAbsolutePositionTab::DocToTxtWriter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExAbsolutePositionTab::DocToTxtWriter::mBuilder", this->mBuilder);

    return result;
}

namespace gtest_test
{

class ExAbsolutePositionTab : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExAbsolutePositionTab> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExAbsolutePositionTab>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExAbsolutePositionTab> ExAbsolutePositionTab::s_instance;

} // namespace gtest_test

void ExAbsolutePositionTab::DocumentToTxt()
{
    // This document contains two sentences separated by an absolute position tab
    auto doc = MakeObject<Document>(MyDir + u"Absolute position tab.docx");

    // An AbsolutePositionTab is a child node of a paragraph
    // AbsolutePositionTabs get picked up when looking for nodes of the SpecialChar type
    SharedPtr<Paragraph> para = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
    auto absPositionTab = System::DynamicCast<Aspose::Words::AbsolutePositionTab>(para->GetChild(Aspose::Words::NodeType::SpecialChar, 0, true));

    // This implementation of the DocumentVisitor pattern converts the document to plain text
    auto myDocToTxtWriter = MakeObject<ExAbsolutePositionTab::DocToTxtWriter>();

    // We can run the DocumentVisitor over the whole first paragraph
    para->Accept(myDocToTxtWriter);

    // A tab character is placed where the AbsolutePositionTab was found
    ASSERT_EQ(u"Before AbsolutePositionTab\tAfter AbsolutePositionTab", myDocToTxtWriter->GetText());

    // An AbsolutePositionTab can accept a DocumentVisitor by itself too
    myDocToTxtWriter = MakeObject<ExAbsolutePositionTab::DocToTxtWriter>();
    absPositionTab->Accept(myDocToTxtWriter);

    ASSERT_EQ(u"\t", myDocToTxtWriter->GetText());
}

namespace gtest_test
{

TEST_F(ExAbsolutePositionTab, DocumentToTxt)
{
    s_instance->DocumentToTxt();
}

} // namespace gtest_test

} // namespace ApiExamples
