// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExAbsolutePositionTab.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3316941060u, ::Aspose::Words::ApiExamples::ExAbsolutePositionTab::DocTextExtractor, ThisTypeBaseTypesInfo);

ExAbsolutePositionTab::DocTextExtractor::DocTextExtractor()
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
}

Aspose::Words::VisitorAction ExAbsolutePositionTab::DocTextExtractor::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    AppendText(run->get_Text());
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExAbsolutePositionTab::DocTextExtractor::VisitAbsolutePositionTab(System::SharedPtr<Aspose::Words::AbsolutePositionTab> tab)
{
    mBuilder->Append(u"\t");
    return Aspose::Words::VisitorAction::Continue;
}

void ExAbsolutePositionTab::DocTextExtractor::AppendText(System::String text)
{
    mBuilder->Append(text);
}

System::String ExAbsolutePositionTab::DocTextExtractor::GetText()
{
    return mBuilder->ToString();
}


RTTI_INFO_IMPL_HASH(435504773u, ::Aspose::Words::ApiExamples::ExAbsolutePositionTab, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExAbsolutePositionTab : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExAbsolutePositionTab> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExAbsolutePositionTab>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExAbsolutePositionTab> ExAbsolutePositionTab::s_instance;

} // namespace gtest_test

void ExAbsolutePositionTab::DocumentToTxt()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Absolute position tab.docx");
    
    // Extract the text contents of our document by accepting this custom document visitor.
    auto myDocTextExtractor = System::MakeObject<Aspose::Words::ApiExamples::ExAbsolutePositionTab::DocTextExtractor>();
    System::SharedPtr<Aspose::Words::Section> fisrtSection = doc->get_FirstSection();
    fisrtSection->get_Body()->Accept(myDocTextExtractor);
    // Visit only start of the document body.
    fisrtSection->get_Body()->AcceptStart(myDocTextExtractor);
    // Visit only end of the document body.
    fisrtSection->get_Body()->AcceptEnd(myDocTextExtractor);
    
    // The absolute position tab, which has no equivalent in string form, has been explicitly converted to a tab character.
    ASSERT_EQ(u"Before AbsolutePositionTab\tAfter AbsolutePositionTab", myDocTextExtractor->GetText());
    
    // An AbsolutePositionTab can accept a DocumentVisitor by itself too.
    auto absPositionTab = System::ExplicitCast<Aspose::Words::AbsolutePositionTab>(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetChild(Aspose::Words::NodeType::SpecialChar, 0, true));
    
    myDocTextExtractor = System::MakeObject<Aspose::Words::ApiExamples::ExAbsolutePositionTab::DocTextExtractor>();
    absPositionTab->Accept(myDocTextExtractor);
    
    ASSERT_EQ(u"\t", myDocTextExtractor->GetText());
}

namespace gtest_test
{

TEST_F(ExAbsolutePositionTab, DocumentToTxt)
{
    s_instance->DocumentToTxt();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
