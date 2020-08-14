// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTxtSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/io/file.h>
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
#include <Aspose.Words.Cpp/Model/Saving/TxtExportHeadersFootersMode.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples {

namespace gtest_test
{

class ExTxtSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExTxtSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExTxtSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExTxtSaveOptions> ExTxtSaveOptions::s_instance;

} // namespace gtest_test

void ExTxtSaveOptions::PageBreaks(bool forcePageBreaks)
{
    //ExStart
    //ExFor:TxtSaveOptionsBase.ForcePageBreaks
    //ExSummary:Shows how to specify whether the page breaks should be preserved during export.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Writeln(u"Page 1");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 3");

    // If ForcePageBreaks is set to true then the output document will have form feed characters in place of page breaks
    // Otherwise, they will be line breaks
    auto saveOptions = MakeObject<TxtSaveOptions>();
    saveOptions->set_ForcePageBreaks(forcePageBreaks);

    doc->Save(ArtifactsDir + u"TxtSaveOptions.PageBreaks.txt", saveOptions);

    // If we load the document using Aspose.Words again, the page breaks will be preserved/lost depending on ForcePageBreaks
    doc = MakeObject<Document>(ArtifactsDir + u"TxtSaveOptions.PageBreaks.txt");

    ASSERT_EQ(forcePageBreaks ? 3 : 1, doc->get_PageCount());
    //ExEnd

    TestUtil::FileContainsString(forcePageBreaks ? String(u"Page 1\r\n\fPage 2\r\n\fPage 3\r\n\r\n") : String(u"Page 1\r\nPage 2\r\nPage 3\r\n\r\n"), ArtifactsDir + u"TxtSaveOptions.PageBreaks.txt");
}

namespace gtest_test
{

using PageBreaks_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::PageBreaks)>::type;

struct PageBreaksVP : public ExTxtSaveOptions, public ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<PageBreaks_Args>
{
    static std::vector<PageBreaks_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(PageBreaksVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PageBreaks(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtSaveOptions, PageBreaksVP, ::testing::ValuesIn(PageBreaksVP::TestCases()));

} // namespace gtest_test

void ExTxtSaveOptions::AddBidiMarks(bool addBidiMarks)
{
    //ExStart
    //ExFor:TxtSaveOptions.AddBidiMarks
    //ExSummary:Shows how to insert Unicode Character 'RIGHT-TO-LEFT MARK' (U+200F) before each bi-directional Run in text.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Writeln(u"Hello world!");
    builder->get_ParagraphFormat()->set_Bidi(true);
    builder->Writeln(u"שלום עולם!");
    builder->Writeln(u"مرحبا بالعالم!");

    auto saveOptions = MakeObject<TxtSaveOptions>();
    saveOptions->set_AddBidiMarks(addBidiMarks);
    saveOptions->set_Encoding(System::Text::Encoding::get_Unicode());

    doc->Save(ArtifactsDir + u"TxtSaveOptions.AddBidiMarks.txt", saveOptions);

    String docText = System::Text::Encoding::get_Unicode()->GetString(System::IO::File::ReadAllBytes(ArtifactsDir + u"TxtSaveOptions.AddBidiMarks.txt"));

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

using AddBidiMarks_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::AddBidiMarks)>::type;

struct AddBidiMarksVP : public ExTxtSaveOptions, public ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<AddBidiMarks_Args>
{
    static std::vector<AddBidiMarks_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(AddBidiMarksVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AddBidiMarks(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtSaveOptions, AddBidiMarksVP, ::testing::ValuesIn(AddBidiMarksVP::TestCases()));

} // namespace gtest_test

void ExTxtSaveOptions::ExportHeadersFooters(Aspose::Words::Saving::TxtExportHeadersFootersMode txtExportHeadersFootersMode)
{
    //ExStart
    //ExFor:TxtSaveOptionsBase.ExportHeadersFootersMode
    //ExFor:TxtExportHeadersFootersMode
    //ExSummary:Shows how to specifies the way headers and footers are exported to plain text format.
    auto doc = MakeObject<Document>();

    // Insert even and primary headers/footers into the document
    // The primary header/footers should override the even ones
    doc->get_FirstSection()->get_HeadersFooters()->Add(MakeObject<HeaderFooter>(doc, Aspose::Words::HeaderFooterType::HeaderEven));
    doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderEven)->AppendParagraph(u"Even header");
    doc->get_FirstSection()->get_HeadersFooters()->Add(MakeObject<HeaderFooter>(doc, Aspose::Words::HeaderFooterType::FooterEven));
    doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterEven)->AppendParagraph(u"Even footer");
    doc->get_FirstSection()->get_HeadersFooters()->Add(MakeObject<HeaderFooter>(doc, Aspose::Words::HeaderFooterType::HeaderPrimary));
    doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->AppendParagraph(u"Primary header");
    doc->get_FirstSection()->get_HeadersFooters()->Add(MakeObject<HeaderFooter>(doc, Aspose::Words::HeaderFooterType::FooterPrimary));
    doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->AppendParagraph(u"Primary footer");

    // Insert pages that would display these headers and footers
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"Page 1");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page 2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Write(u"Page 3");

    // Three values are available in TxtExportHeadersFootersMode enum:
    // "None" - No headers and footers are exported
    // "AllAtEnd" - All headers and footers are placed after all section bodies at the very end of a document
    // "PrimaryOnly" - Only primary headers and footers are exported at the beginning and end of each section (default value)
    auto saveOptions = MakeObject<TxtSaveOptions>();
    saveOptions->set_ExportHeadersFootersMode(txtExportHeadersFootersMode);

    doc->Save(ArtifactsDir + u"TxtSaveOptions.ExportHeadersFooters.txt", saveOptions);

    String docText = System::IO::File::ReadAllText(ArtifactsDir + u"TxtSaveOptions.ExportHeadersFooters.txt");

    switch (txtExportHeadersFootersMode)
    {
        case Aspose::Words::Saving::TxtExportHeadersFootersMode::AllAtEnd:
            ASSERT_EQ(String(u"Page 1\r\n") + u"Page 2\r\n" + u"Page 3\r\n" + u"Even header\r\n\r\n" + u"Primary header\r\n\r\n" + u"Even footer\r\n\r\n" + u"Primary footer\r\n\r\n", docText);
            break;

        case Aspose::Words::Saving::TxtExportHeadersFootersMode::PrimaryOnly:
            ASSERT_EQ(String(u"Primary header\r\n") + u"Page 1\r\n" + u"Page 2\r\n" + u"Page 3\r\n" + u"Primary footer\r\n", docText);
            break;

        case Aspose::Words::Saving::TxtExportHeadersFootersMode::None:
            ASSERT_EQ(String(u"Page 1\r\n") + u"Page 2\r\n" + u"Page 3\r\n", docText);
            break;

    }
    //ExEnd
}

namespace gtest_test
{

using ExportHeadersFooters_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::ExportHeadersFooters)>::type;

struct ExportHeadersFootersVP : public ExTxtSaveOptions, public ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<ExportHeadersFooters_Args>
{
    static std::vector<ExportHeadersFooters_Args> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::TxtExportHeadersFootersMode::AllAtEnd),
            std::make_tuple(Aspose::Words::Saving::TxtExportHeadersFootersMode::PrimaryOnly),
            std::make_tuple(Aspose::Words::Saving::TxtExportHeadersFootersMode::None),
        };
    }
};

TEST_P(ExportHeadersFootersVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportHeadersFooters(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtSaveOptions, ExportHeadersFootersVP, ::testing::ValuesIn(ExportHeadersFootersVP::TestCases()));

} // namespace gtest_test

void ExTxtSaveOptions::TxtListIndentation()
{
    //ExStart
    //ExFor:TxtListIndentation
    //ExFor:TxtListIndentation.Count
    //ExFor:TxtListIndentation.Character
    //ExFor:TxtSaveOptions.ListIndentation
    //ExSummary:Shows how to configure list indenting when converting to plaintext.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Create a list with three levels of indentation
    builder->get_ListFormat()->ApplyNumberDefault();
    builder->Writeln(u"Item 1");
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"Item 2");
    builder->get_ListFormat()->ListIndent();
    builder->Write(u"Item 3");

    // Microsoft Word list objects get lost when converting to plaintext
    // We can create a custom representation for list indentation using pure plaintext with a SaveOptions object
    // In this case, each list item will be left-padded by 3 space characters times its list indent level
    auto txtSaveOptions = MakeObject<TxtSaveOptions>();
    txtSaveOptions->get_ListIndentation()->set_Count(3);
    txtSaveOptions->get_ListIndentation()->set_Character(u' ');

    doc->Save(ArtifactsDir + u"TxtSaveOptions.TxtListIndentation.txt", txtSaveOptions);

    String docText = System::IO::File::ReadAllText(ArtifactsDir + u"TxtSaveOptions.TxtListIndentation.txt");

    ASSERT_EQ(String(u"1. Item 1\r\n") + u"   a. Item 2\r\n" + u"      i. Item 3\r\n", docText);
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
    //ExSummary:Shows how to change the appearance of lists when converting to plaintext.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Create a bulleted list with five levels of indentation
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

    // The SimplifyListLabels flag will convert some list symbols
    // into ASCII characters such as *, o, +, > etc, depending on list level
    auto txtSaveOptions = MakeObject<TxtSaveOptions>();
    txtSaveOptions->set_SimplifyListLabels(simplifyListLabels);

    doc->Save(ArtifactsDir + u"TxtSaveOptions.SimplifyListLabels.txt", txtSaveOptions);

    String docText = System::IO::File::ReadAllText(ArtifactsDir + u"TxtSaveOptions.SimplifyListLabels.txt");

    if (simplifyListLabels)
    {
        ASSERT_EQ(String(u"* Item 1\r\n") + u"  > Item 2\r\n" + u"    + Item 3\r\n" + u"      - Item 4\r\n" + u"        o Item 5\r\n", docText);
    }
    else
    {
        ASSERT_EQ(String(u"· Item 1\r\n") + u"o Item 2\r\n" + u"§ Item 3\r\n" + u"· Item 4\r\n" + u"o Item 5\r\n", docText);
    }
    //ExEnd
}

namespace gtest_test
{

using SimplifyListLabels_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::SimplifyListLabels)>::type;

struct SimplifyListLabelsVP : public ExTxtSaveOptions, public ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<SimplifyListLabels_Args>
{
    static std::vector<SimplifyListLabels_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(SimplifyListLabelsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SimplifyListLabels(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtSaveOptions, SimplifyListLabelsVP, ::testing::ValuesIn(SimplifyListLabelsVP::TestCases()));

} // namespace gtest_test

void ExTxtSaveOptions::ParagraphBreak()
{
    //ExStart
    //ExFor:TxtSaveOptions
    //ExFor:TxtSaveOptions.SaveFormat
    //ExFor:TxtSaveOptionsBase
    //ExFor:TxtSaveOptionsBase.ParagraphBreak
    //ExSummary:Shows how to save a .txt document with a custom paragraph break.
    // Create a new document and add some paragraphs
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Writeln(u"Paragraph 1.");
    builder->Writeln(u"Paragraph 2.");
    builder->Write(u"Paragraph 3.");

    // When saved to plain text, the paragraphs we created can be separated by a custom string
    auto txtSaveOptions = MakeObject<TxtSaveOptions>();
    txtSaveOptions->set_SaveFormat(Aspose::Words::SaveFormat::Text);
    txtSaveOptions->set_ParagraphBreak(u" End of paragraph.\n\n\t");

    doc->Save(ArtifactsDir + u"TxtSaveOptions.ParagraphBreak.txt", txtSaveOptions);

    String docText = System::IO::File::ReadAllText(ArtifactsDir + u"TxtSaveOptions.ParagraphBreak.txt");

    ASSERT_EQ(String(u"Paragraph 1. End of paragraph.\n\n\t") + u"Paragraph 2. End of paragraph.\n\n\t" + u"Paragraph 3. End of paragraph.\n\n\t", docText);
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
    // Create a new document and add some text from outside the ASCII character set
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Write(u"À È Ì Ò Ù.");

    // We can use a SaveOptions object to make sure the encoding we save the .txt document in supports our content
    auto txtSaveOptions = MakeObject<TxtSaveOptions>();
    txtSaveOptions->set_Encoding(System::Text::Encoding::get_UTF8());

    doc->Save(ArtifactsDir + u"TxtSaveOptions.Encoding.txt", txtSaveOptions);

    String docText = System::Text::Encoding::get_UTF8()->GetString(System::IO::File::ReadAllBytes(ArtifactsDir + u"TxtSaveOptions.Encoding.txt"));

    ASSERT_EQ(u"\ufeffÀ È Ì Ò Ù.\r\n", docText);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTxtSaveOptions, Encoding)
{
    s_instance->Encoding();
}

} // namespace gtest_test

void ExTxtSaveOptions::TableLayout(bool preserveTableLayout)
{
    //ExStart
    //ExFor:TxtSaveOptions.PreserveTableLayout
    //ExSummary:Shows how to preserve the layout of tables when converting to plaintext.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a table
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

    // Tables, with their borders and widths do not translate to plaintext
    // However, we can configure a SaveOptions object to arrange table contents to preserve some of the table's appearance
    auto txtSaveOptions = MakeObject<TxtSaveOptions>();
    txtSaveOptions->set_PreserveTableLayout(preserveTableLayout);

    doc->Save(ArtifactsDir + u"TxtSaveOptions.TableLayout.txt", txtSaveOptions);

    String docText = System::IO::File::ReadAllText(ArtifactsDir + u"TxtSaveOptions.TableLayout.txt");

    if (preserveTableLayout)
    {
        ASSERT_EQ(String(u"Row 1, cell 1                Row 1, cell 2\r\n") + u"Row 2, cell 1                Row 2, cell 2\r\n\r\n", docText);
    }
    else
    {
        ASSERT_EQ(String(u"Row 1, cell 1\r\n") + u"Row 1, cell 2\r\n" + u"Row 2, cell 1\r\n" + u"Row 2, cell 2\r\n\r\n", docText);
    }
    //ExEnd
}

namespace gtest_test
{

using TableLayout_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::TableLayout)>::type;

struct TableLayoutVP : public ExTxtSaveOptions, public ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<TableLayout_Args>
{
    static std::vector<TableLayout_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(TableLayoutVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TableLayout(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtSaveOptions, TableLayoutVP, ::testing::ValuesIn(TableLayoutVP::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
