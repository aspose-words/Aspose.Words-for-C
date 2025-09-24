// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTxtSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/io/file.h>
#include <system/environment.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Saving/TxtSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/TxtListIndentation.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

#include "TestUtil.h"


using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3296663617u, ::Aspose::Words::ApiExamples::ExTxtSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExTxtSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExTxtSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExTxtSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExTxtSaveOptions> ExTxtSaveOptions::s_instance;

} // namespace gtest_test

void ExTxtSaveOptions::PageBreaks(bool forcePageBreaks)
{
    //ExStart
    //ExFor:TxtSaveOptionsBase.ForcePageBreaks
    //ExSummary:Shows how to specify whether to preserve page breaks when exporting a document to plaintext.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Page 1");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 3");
    
    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save"
    // method to modify how we save the document to plaintext.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::TxtSaveOptions>();
    
    // The Aspose.Words "Document" objects have page breaks, just like Microsoft Word documents.
    // Save formats such as ".txt" are one continuous body of text without page breaks.
    // Set the "ForcePageBreaks" property to "true" to preserve all page breaks in the form of '\f' characters.
    // Set the "ForcePageBreaks" property to "false" to discard all page breaks.
    saveOptions->set_ForcePageBreaks(forcePageBreaks);
    
    doc->Save(get_ArtifactsDir() + u"TxtSaveOptions.PageBreaks.txt", saveOptions);
    
    // If we load a plaintext document with page breaks,
    // the "Document" object will use them to split the body into pages.
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"TxtSaveOptions.PageBreaks.txt");
    
    ASSERT_EQ(forcePageBreaks ? 3 : 1, doc->get_PageCount());
    //ExEnd
    
    Aspose::Words::ApiExamples::TestUtil::FileContainsString(forcePageBreaks ? System::String(u"Page 1\r\n\fPage 2\r\n\fPage 3\r\n\r\n") : System::String(u"Page 1\r\nPage 2\r\nPage 3\r\n\r\n"), get_ArtifactsDir() + u"TxtSaveOptions.PageBreaks.txt");
}

namespace gtest_test
{

using ExTxtSaveOptions_PageBreaks_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExTxtSaveOptions::PageBreaks)>::type;

struct ExTxtSaveOptions_PageBreaks : public ExTxtSaveOptions, public Aspose::Words::ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<ExTxtSaveOptions_PageBreaks_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExTxtSaveOptions_PageBreaks, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PageBreaks(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtSaveOptions_PageBreaks, ::testing::ValuesIn(ExTxtSaveOptions_PageBreaks::TestCases()));

} // namespace gtest_test

void ExTxtSaveOptions::AddBidiMarks(bool addBidiMarks)
{
    //ExStart
    //ExFor:TxtSaveOptions.AddBidiMarks
    //ExSummary:Shows how to insert Unicode Character 'RIGHT-TO-LEFT MARK' (U+200F) before each bi-directional Run in text.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Hello world!");
    builder->get_ParagraphFormat()->set_Bidi(true);
    builder->Writeln(u"שלום עולם!");
    builder->Writeln(u"مرحبا بالعالم!");
    
    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::TxtSaveOptions>();
    saveOptions->set_Encoding(System::Text::Encoding::get_Unicode());
    
    // Set the "AddBidiMarks" property to "true" to add marks before runs
    // with right-to-left text to indicate the fact.
    // Set the "AddBidiMarks" property to "false" to write all left-to-right
    // and right-to-left run equally with nothing to indicate which is which.
    saveOptions->set_AddBidiMarks(addBidiMarks);
    
    doc->Save(get_ArtifactsDir() + u"TxtSaveOptions.AddBidiMarks.txt", saveOptions);
    
    System::String docText = System::Text::Encoding::get_Unicode()->GetString(System::IO::File::ReadAllBytes(get_ArtifactsDir() + u"TxtSaveOptions.AddBidiMarks.txt"));
    
    if (addBidiMarks)
    {
        ASSERT_EQ(u"\ufeffHello world!‎\r\nשלום עולם!‏\r\nمرحبا بالعالم!‏\r\n\r\n", docText);
        ASSERT_TRUE(docText.Contains(u"\u200f"));
    }
    else
    {
        ASSERT_EQ(u"\ufeffHello world!\r\nשלום עולם!\r\nمرحبا بالعالم!\r\n\r\n", docText);
        ASSERT_FALSE(docText.Contains(u"\u200f"));
    }
    //ExEnd
}

namespace gtest_test
{

using ExTxtSaveOptions_AddBidiMarks_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExTxtSaveOptions::AddBidiMarks)>::type;

struct ExTxtSaveOptions_AddBidiMarks : public ExTxtSaveOptions, public Aspose::Words::ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<ExTxtSaveOptions_AddBidiMarks_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExTxtSaveOptions_AddBidiMarks, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AddBidiMarks(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtSaveOptions_AddBidiMarks, ::testing::ValuesIn(ExTxtSaveOptions_AddBidiMarks::TestCases()));

} // namespace gtest_test

void ExTxtSaveOptions::ExportHeadersFooters(Aspose::Words::Saving::TxtExportHeadersFootersMode txtExportHeadersFootersMode)
{
    //ExStart
    //ExFor:TxtSaveOptionsBase.ExportHeadersFootersMode
    //ExFor:TxtExportHeadersFootersMode
    //ExSummary:Shows how to specify how to export headers and footers to plain text format.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Insert even and primary headers/footers into the document.
    // The primary header/footers will override the even headers/footers.
    doc->get_FirstSection()->get_HeadersFooters()->Add(System::MakeObject<Aspose::Words::HeaderFooter>(doc, Aspose::Words::HeaderFooterType::HeaderEven));
    doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderEven)->AppendParagraph(u"Even header");
    doc->get_FirstSection()->get_HeadersFooters()->Add(System::MakeObject<Aspose::Words::HeaderFooter>(doc, Aspose::Words::HeaderFooterType::FooterEven));
    doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterEven)->AppendParagraph(u"Even footer");
    doc->get_FirstSection()->get_HeadersFooters()->Add(System::MakeObject<Aspose::Words::HeaderFooter>(doc, Aspose::Words::HeaderFooterType::HeaderPrimary));
    doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->AppendParagraph(u"Primary header");
    doc->get_FirstSection()->get_HeadersFooters()->Add(System::MakeObject<Aspose::Words::HeaderFooter>(doc, Aspose::Words::HeaderFooterType::FooterPrimary));
    doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->AppendParagraph(u"Primary footer");
    
    // Insert pages to display these headers and footers.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Page 1");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Write(u"Page 3");
    
    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::TxtSaveOptions>();
    
    // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.None"
    // to not export any headers/footers.
    // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.PrimaryOnly"
    // to only export primary headers/footers.
    // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.AllAtEnd"
    // to place all headers and footers for all section bodies at the end of the document.
    saveOptions->set_ExportHeadersFootersMode(txtExportHeadersFootersMode);
    
    doc->Save(get_ArtifactsDir() + u"TxtSaveOptions.ExportHeadersFooters.txt", saveOptions);
    
    System::String docText = System::IO::File::ReadAllText(get_ArtifactsDir() + u"TxtSaveOptions.ExportHeadersFooters.txt");
    
    System::String newLine = System::Environment::get_NewLine();
    switch (txtExportHeadersFootersMode)
    {
        case Aspose::Words::Saving::TxtExportHeadersFootersMode::AllAtEnd:
            ASSERT_EQ(System::String::Format(u"Page 1{0}", newLine) + System::String::Format(u"Page 2{0}", newLine) + System::String::Format(u"Page 3{0}", newLine) + System::String::Format(u"Even header{0}{1}", newLine, newLine) + System::String::Format(u"Primary header{0}{1}", newLine, newLine) + System::String::Format(u"Even footer{0}{1}", newLine, newLine) + System::String::Format(u"Primary footer{0}{1}", newLine, newLine), docText);
            break;
        
        case Aspose::Words::Saving::TxtExportHeadersFootersMode::PrimaryOnly:
            ASSERT_EQ(System::String::Format(u"Primary header{0}", newLine) + System::String::Format(u"Page 1{0}", newLine) + System::String::Format(u"Page 2{0}", newLine) + System::String::Format(u"Page 3{0}", newLine) + System::String::Format(u"Primary footer{0}", newLine), docText);
            break;
        
        case Aspose::Words::Saving::TxtExportHeadersFootersMode::None:
            ASSERT_EQ(System::String::Format(u"Page 1{0}", newLine) + System::String::Format(u"Page 2{0}", newLine) + System::String::Format(u"Page 3{0}", newLine), docText);
            break;
        
    }
    //ExEnd
}

namespace gtest_test
{

using ExTxtSaveOptions_ExportHeadersFooters_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExTxtSaveOptions::ExportHeadersFooters)>::type;

struct ExTxtSaveOptions_ExportHeadersFooters : public ExTxtSaveOptions, public Aspose::Words::ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<ExTxtSaveOptions_ExportHeadersFooters_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::TxtExportHeadersFootersMode::AllAtEnd),
            std::make_tuple(Aspose::Words::Saving::TxtExportHeadersFootersMode::PrimaryOnly),
            std::make_tuple(Aspose::Words::Saving::TxtExportHeadersFootersMode::None),
        };
    }
};

TEST_P(ExTxtSaveOptions_ExportHeadersFooters, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportHeadersFooters(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtSaveOptions_ExportHeadersFooters, ::testing::ValuesIn(ExTxtSaveOptions_ExportHeadersFooters::TestCases()));

} // namespace gtest_test

void ExTxtSaveOptions::TxtListIndentation()
{
    //ExStart
    //ExFor:TxtListIndentation
    //ExFor:TxtListIndentation.Count
    //ExFor:TxtListIndentation.Character
    //ExFor:TxtSaveOptions.ListIndentation
    //ExSummary:Shows how to configure list indenting when saving a document to plaintext.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create a list with three levels of indentation.
    builder->get_ListFormat()->ApplyNumberDefault();
    builder->Writeln(u"Item 1");
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"Item 2");
    builder->get_ListFormat()->ListIndent();
    builder->Write(u"Item 3");
    
    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    auto txtSaveOptions = System::MakeObject<Aspose::Words::Saving::TxtSaveOptions>();
    
    // Set the "Character" property to assign a character to use
    // for padding that simulates list indentation in plaintext.
    txtSaveOptions->get_ListIndentation()->set_Character(u' ');
    
    // Set the "Count" property to specify the number of times
    // to place the padding character for each list indent level.
    txtSaveOptions->get_ListIndentation()->set_Count(3);
    
    doc->Save(get_ArtifactsDir() + u"TxtSaveOptions.TxtListIndentation.txt", txtSaveOptions);
    
    System::String docText = System::IO::File::ReadAllText(get_ArtifactsDir() + u"TxtSaveOptions.TxtListIndentation.txt");
    System::String newLine = System::Environment::get_NewLine();
    
    ASSERT_EQ(System::String::Format(u"1. Item 1{0}", newLine) + System::String::Format(u"   a. Item 2{0}", newLine) + System::String::Format(u"      i. Item 3{0}", newLine), docText);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTxtSaveOptions, TxtListIndentation)
{
    s_instance->TxtListIndentation();
}

} // namespace gtest_test

void ExTxtSaveOptions::SimplifyListLabels(bool simplifyListLabels)
{
    //ExStart
    //ExFor:TxtSaveOptions.SimplifyListLabels
    //ExSummary:Shows how to change the appearance of lists when saving a document to plaintext.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create a bulleted list with five levels of indentation.
    builder->get_ListFormat()->ApplyBulletDefault();
    builder->Writeln(u"Item 1");
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"Item 2");
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"Item 3");
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"Item 4");
    builder->get_ListFormat()->ListIndent();
    builder->Write(u"Item 5");
    
    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    auto txtSaveOptions = System::MakeObject<Aspose::Words::Saving::TxtSaveOptions>();
    
    // Set the "SimplifyListLabels" property to "true" to convert some list
    // symbols into simpler ASCII characters, such as '*', 'o', '+', '>', etc.
    // Set the "SimplifyListLabels" property to "false" to preserve as many original list symbols as possible.
    txtSaveOptions->set_SimplifyListLabels(simplifyListLabels);
    
    doc->Save(get_ArtifactsDir() + u"TxtSaveOptions.SimplifyListLabels.txt", txtSaveOptions);
    
    System::String docText = System::IO::File::ReadAllText(get_ArtifactsDir() + u"TxtSaveOptions.SimplifyListLabels.txt");
    
    System::String newLine = System::Environment::get_NewLine();
    if (simplifyListLabels)
    {
        ASSERT_EQ(System::String::Format(u"* Item 1{0}", newLine) + System::String::Format(u"  > Item 2{0}", newLine) + System::String::Format(u"    + Item 3{0}", newLine) + System::String::Format(u"      - Item 4{0}", newLine) + System::String::Format(u"        o Item 5{0}", newLine), docText);
    }
    else
    {
        ASSERT_EQ(System::String::Format(u"· Item 1{0}", newLine) + System::String::Format(u"o Item 2{0}", newLine) + System::String::Format(u"§ Item 3{0}", newLine) + System::String::Format(u"· Item 4{0}", newLine) + System::String::Format(u"o Item 5{0}", newLine), docText);
    }
    //ExEnd
}

namespace gtest_test
{

using ExTxtSaveOptions_SimplifyListLabels_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExTxtSaveOptions::SimplifyListLabels)>::type;

struct ExTxtSaveOptions_SimplifyListLabels : public ExTxtSaveOptions, public Aspose::Words::ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<ExTxtSaveOptions_SimplifyListLabels_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExTxtSaveOptions_SimplifyListLabels, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SimplifyListLabels(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtSaveOptions_SimplifyListLabels, ::testing::ValuesIn(ExTxtSaveOptions_SimplifyListLabels::TestCases()));

} // namespace gtest_test

void ExTxtSaveOptions::ParagraphBreak()
{
    //ExStart
    //ExFor:TxtSaveOptions
    //ExFor:TxtSaveOptions.SaveFormat
    //ExFor:TxtSaveOptionsBase
    //ExFor:TxtSaveOptionsBase.ParagraphBreak
    //ExSummary:Shows how to save a .txt document with a custom paragraph break.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Paragraph 1.");
    builder->Writeln(u"Paragraph 2.");
    builder->Write(u"Paragraph 3.");
    
    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    auto txtSaveOptions = System::MakeObject<Aspose::Words::Saving::TxtSaveOptions>();
    
    ASSERT_EQ(Aspose::Words::SaveFormat::Text, txtSaveOptions->get_SaveFormat());
    
    // Set the "ParagraphBreak" to a custom value that we wish to put at the end of every paragraph.
    txtSaveOptions->set_ParagraphBreak(u" End of paragraph.\n\n\t");
    
    doc->Save(get_ArtifactsDir() + u"TxtSaveOptions.ParagraphBreak.txt", txtSaveOptions);
    
    System::String docText = System::IO::File::ReadAllText(get_ArtifactsDir() + u"TxtSaveOptions.ParagraphBreak.txt");
    
    ASSERT_EQ(System::String(u"Paragraph 1. End of paragraph.\n\n\t") + u"Paragraph 2. End of paragraph.\n\n\t" + u"Paragraph 3. End of paragraph.\n\n\t", docText);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTxtSaveOptions, ParagraphBreak)
{
    s_instance->ParagraphBreak();
}

} // namespace gtest_test

void ExTxtSaveOptions::Encoding()
{
    //ExStart
    //ExFor:TxtSaveOptionsBase.Encoding
    //ExSummary:Shows how to set encoding for a .txt output document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Add some text with characters from outside the ASCII character set.
    builder->Write(u"À È Ì Ò Ù.");
    
    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    auto txtSaveOptions = System::MakeObject<Aspose::Words::Saving::TxtSaveOptions>();
    
    // Verify that the "Encoding" property contains the appropriate encoding for our document's contents.
    ASPOSE_ASSERT_EQ(System::Text::Encoding::get_UTF8(), txtSaveOptions->get_Encoding());
    
    doc->Save(get_ArtifactsDir() + u"TxtSaveOptions.Encoding.UTF8.txt", txtSaveOptions);
    
    System::String docText = System::Text::Encoding::get_UTF8()->GetString(System::IO::File::ReadAllBytes(get_ArtifactsDir() + u"TxtSaveOptions.Encoding.UTF8.txt"));
    
    ASSERT_EQ(u"\ufeffÀ È Ì Ò Ù.\r\n", docText);
    
    // Using an unsuitable encoding may result in a loss of document contents.
    txtSaveOptions->set_Encoding(System::Text::Encoding::get_ASCII());
    doc->Save(get_ArtifactsDir() + u"TxtSaveOptions.Encoding.ASCII.txt", txtSaveOptions);
    docText = System::Text::Encoding::get_ASCII()->GetString(System::IO::File::ReadAllBytes(get_ArtifactsDir() + u"TxtSaveOptions.Encoding.ASCII.txt"));
    
    ASSERT_EQ(u"? ? ? ? ?.\r\n", docText);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTxtSaveOptions, Encoding)
{
    s_instance->Encoding();
}

} // namespace gtest_test

void ExTxtSaveOptions::PreserveTableLayout(bool preserveTableLayout)
{
    //ExStart
    //ExFor:TxtSaveOptions.PreserveTableLayout
    //ExSummary:Shows how to preserve the layout of tables when converting to plaintext.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Row 1, cell 1");
    builder->InsertCell();
    builder->Write(u"Row 1, cell 2");
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"Row 2, cell 1");
    builder->InsertCell();
    builder->Write(u"Row 2, cell 2");
    builder->EndTable();
    
    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    auto txtSaveOptions = System::MakeObject<Aspose::Words::Saving::TxtSaveOptions>();
    
    // Set the "PreserveTableLayout" property to "true" to apply whitespace padding to the contents
    // of the output plaintext document to preserve as much of the table's layout as possible.
    // Set the "PreserveTableLayout" property to "false" to save all tables' contents
    // as a continuous body of text, with just a new line for each row.
    txtSaveOptions->set_PreserveTableLayout(preserveTableLayout);
    
    doc->Save(get_ArtifactsDir() + u"TxtSaveOptions.PreserveTableLayout.txt", txtSaveOptions);
    
    System::String docText = System::IO::File::ReadAllText(get_ArtifactsDir() + u"TxtSaveOptions.PreserveTableLayout.txt");
    
    if (preserveTableLayout)
    {
        ASSERT_EQ(System::String(u"Row 1, cell 1                                            Row 1, cell 2\r\n") + u"Row 2, cell 1                                            Row 2, cell 2\r\n\r\n", docText);
    }
    else
    {
        ASSERT_EQ(System::String(u"Row 1, cell 1\r") + u"Row 1, cell 2\r" + u"Row 2, cell 1\r" + u"Row 2, cell 2\r\r\n", docText);
    }
    //ExEnd
}

namespace gtest_test
{

using ExTxtSaveOptions_PreserveTableLayout_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExTxtSaveOptions::PreserveTableLayout)>::type;

struct ExTxtSaveOptions_PreserveTableLayout : public ExTxtSaveOptions, public Aspose::Words::ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<ExTxtSaveOptions_PreserveTableLayout_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExTxtSaveOptions_PreserveTableLayout, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PreserveTableLayout(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtSaveOptions_PreserveTableLayout, ::testing::ValuesIn(ExTxtSaveOptions_PreserveTableLayout::TestCases()));

} // namespace gtest_test

void ExTxtSaveOptions::MaxCharactersPerLine()
{
    //ExStart
    //ExFor:TxtSaveOptions.MaxCharactersPerLine
    //ExSummary:Shows how to set maximum number of characters per line.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(System::String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ") + u"Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");
    
    // Set 30 characters as maximum allowed per one line.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::TxtSaveOptions>();
    saveOptions->set_MaxCharactersPerLine(30);
    
    doc->Save(get_ArtifactsDir() + u"TxtSaveOptions.MaxCharactersPerLine.txt", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTxtSaveOptions, MaxCharactersPerLine)
{
    s_instance->MaxCharactersPerLine();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
