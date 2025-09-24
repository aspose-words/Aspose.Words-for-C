// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExMarkdownLoadOptions.h"

#include <testing/test_predicates.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/io/memory_stream.h>
#include <system/environment.h>
#include <system/details/dispose_guard.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Underline.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Loading/MarkdownLoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Loading;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(708566009u, ::Aspose::Words::ApiExamples::ExMarkdownLoadOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExMarkdownLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExMarkdownLoadOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExMarkdownLoadOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExMarkdownLoadOptions> ExMarkdownLoadOptions::s_instance;

} // namespace gtest_test

void ExMarkdownLoadOptions::PreserveEmptyLines()
{
    //ExStart:PreserveEmptyLines
    //GistId:a775441ecb396eea917a2717cb9e8f8f
    //ExFor:MarkdownLoadOptions
    //ExFor:MarkdownLoadOptions.#ctor
    //ExFor:MarkdownLoadOptions.PreserveEmptyLines
    //ExSummary:Shows how to preserve empty line while load a document.
    System::String mdText = System::String::Format(u"{0}Line1{1}{2}Line2{3}{4}", System::Environment::get_NewLine(), System::Environment::get_NewLine(), System::Environment::get_NewLine(), System::Environment::get_NewLine(), System::Environment::get_NewLine());
    {
        auto stream = System::MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(mdText));
        auto loadOptions = System::MakeObject<Aspose::Words::Loading::MarkdownLoadOptions>();
        loadOptions->set_PreserveEmptyLines(true);
        auto doc = System::MakeObject<Aspose::Words::Document>(stream, loadOptions);
        
        ASSERT_EQ(u"\rLine1\r\rLine2\r\f", doc->GetText());
    }
    //ExEnd:PreserveEmptyLines
}

namespace gtest_test
{

TEST_F(ExMarkdownLoadOptions, PreserveEmptyLines)
{
    s_instance->PreserveEmptyLines();
}

} // namespace gtest_test

void ExMarkdownLoadOptions::ImportUnderlineFormatting()
{
    //ExStart:ImportUnderlineFormatting
    //GistId:e06aa7a168b57907a5598e823a22bf0a
    //ExFor:MarkdownLoadOptions.ImportUnderlineFormatting
    //ExSummary:Shows how to recognize plus characters "++" as underline text formatting.
    {
        auto stream = System::MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_ASCII()->GetBytes(u"++12 and B++"));
        auto loadOptions = System::MakeObject<Aspose::Words::Loading::MarkdownLoadOptions>();
        loadOptions->set_ImportUnderlineFormatting(true);
        auto doc = System::MakeObject<Aspose::Words::Document>(stream, loadOptions);
        
        auto para = System::ExplicitCast<Aspose::Words::Paragraph>(doc->GetChild(Aspose::Words::NodeType::Paragraph, 0, true));
        ASSERT_EQ(Aspose::Words::Underline::Single, para->get_Runs()->idx_get(0)->get_Font()->get_Underline());
        
        loadOptions = System::MakeObject<Aspose::Words::Loading::MarkdownLoadOptions>();
        loadOptions->set_ImportUnderlineFormatting(false);
        doc = System::MakeObject<Aspose::Words::Document>(stream, loadOptions);
        
        para = System::ExplicitCast<Aspose::Words::Paragraph>(doc->GetChild(Aspose::Words::NodeType::Paragraph, 0, true));
        ASSERT_EQ(Aspose::Words::Underline::None, para->get_Runs()->idx_get(0)->get_Font()->get_Underline());
    }
    //ExEnd:ImportUnderlineFormatting
}

namespace gtest_test
{

TEST_F(ExMarkdownLoadOptions, ImportUnderlineFormatting)
{
    s_instance->ImportUnderlineFormatting();
}

} // namespace gtest_test

void ExMarkdownLoadOptions::SoftLineBreakCharacter()
{
    //ExStart:SoftLineBreakCharacter
    //GistId:571cc6e23284a2ec075d15d4c32e3bbf
    //ExFor:MarkdownLoadOptions.SoftLineBreakCharacter
    //ExSummary:Shows how to set soft line break character.
    {
        auto stream = System::MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(u"line1\nline2"));
        auto loadOptions = System::MakeObject<Aspose::Words::Loading::MarkdownLoadOptions>();
        loadOptions->set_SoftLineBreakCharacter(Aspose::Words::ControlChar::LineBreakChar);
        auto doc = System::MakeObject<Aspose::Words::Document>(stream, loadOptions);
        
        ASSERT_EQ(u"line1\u000bline2", doc->GetText().Trim());
    }
    //ExEnd:SoftLineBreakCharacter
}

namespace gtest_test
{

TEST_F(ExMarkdownLoadOptions, SoftLineBreakCharacter)
{
    s_instance->SoftLineBreakCharacter();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
