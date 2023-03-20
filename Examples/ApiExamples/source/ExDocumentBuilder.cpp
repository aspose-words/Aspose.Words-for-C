// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentBuilder.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::DigitalSignatures;
namespace ApiExamples { namespace gtest_test {

class ExDocumentBuilder : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExDocumentBuilder> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExDocumentBuilder>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExDocumentBuilder> ExDocumentBuilder::s_instance;

TEST_F(ExDocumentBuilder, WriteAndFont)
{
    s_instance->WriteAndFont();
}

TEST_F(ExDocumentBuilder, HeadersAndFooters)
{
    s_instance->HeadersAndFooters();
}

TEST_F(ExDocumentBuilder, MergeFields)
{
    s_instance->MergeFields();
}

TEST_F(ExDocumentBuilder, InsertHorizontalRule)
{
    s_instance->InsertHorizontalRule();
}

TEST_F(ExDocumentBuilder, HorizontalRuleFormatExceptions)
{
    RecordProperty("Description", "Checking the boundary conditions of WidthPercent and Height properties");

    s_instance->HorizontalRuleFormatExceptions();
}

TEST_F(ExDocumentBuilder, InsertHyperlink)
{
    s_instance->InsertHyperlink();
}

TEST_F(ExDocumentBuilder, PushPopFont)
{
    s_instance->PushPopFont();
}

TEST_F(ExDocumentBuilder, InsertWatermark)
{
    s_instance->InsertWatermark();
}

TEST_F(ExDocumentBuilder, InsertHtml)
{
    s_instance->InsertHtml();
}

using ExDocumentBuilder_InsertHtmlWithFormatting_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::InsertHtmlWithFormatting)>::type;

struct ExDocumentBuilder_InsertHtmlWithFormatting : public ExDocumentBuilder,
                                                    public ApiExamples::ExDocumentBuilder,
                                                    public ::testing::WithParamInterface<ExDocumentBuilder_InsertHtmlWithFormatting_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocumentBuilder_InsertHtmlWithFormatting, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->InsertHtmlWithFormatting(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_InsertHtmlWithFormatting, ::testing::ValuesIn(ExDocumentBuilder_InsertHtmlWithFormatting::TestCases()));

TEST_F(ExDocumentBuilder, MathMl)
{
    s_instance->MathMl();
}

TEST_F(ExDocumentBuilder, InsertTextAndBookmark)
{
    s_instance->InsertTextAndBookmark();
}

TEST_F(ExDocumentBuilder, CreateColumnBookmark)
{
    s_instance->CreateColumnBookmark();
}

TEST_F(ExDocumentBuilder, CreateForm)
{
    s_instance->CreateForm();
}

TEST_F(ExDocumentBuilder, InsertCheckBox)
{
    s_instance->InsertCheckBox();
}

TEST_F(ExDocumentBuilder, InsertCheckBoxEmptyName)
{
    s_instance->InsertCheckBoxEmptyName();
}

TEST_F(ExDocumentBuilder, WorkingWithNodes)
{
    s_instance->WorkingWithNodes();
}

TEST_F(ExDocumentBuilder, FillMergeFields)
{
    s_instance->FillMergeFields();
}

TEST_F(ExDocumentBuilder, InsertToc)
{
    s_instance->InsertToc();
}

TEST_F(ExDocumentBuilder, InsertTable)
{
    s_instance->InsertTable();
}

TEST_F(ExDocumentBuilder, InsertTableWithStyle)
{
    s_instance->InsertTableWithStyle();
}

TEST_F(ExDocumentBuilder, InsertTableSetHeadingRow)
{
    s_instance->InsertTableSetHeadingRow();
}

TEST_F(ExDocumentBuilder, InsertTableWithPreferredWidth)
{
    s_instance->InsertTableWithPreferredWidth();
}

TEST_F(ExDocumentBuilder, InsertCellsWithPreferredWidths)
{
    s_instance->InsertCellsWithPreferredWidths();
}

TEST_F(ExDocumentBuilder, InsertTableFromHtml)
{
    s_instance->InsertTableFromHtml();
}

TEST_F(ExDocumentBuilder, InsertNestedTable)
{
    s_instance->InsertNestedTable();
}

TEST_F(ExDocumentBuilder, CreateTable)
{
    s_instance->CreateTable();
}

TEST_F(ExDocumentBuilder, BuildFormattedTable)
{
    s_instance->BuildFormattedTable();
}

TEST_F(ExDocumentBuilder, TableBordersAndShading)
{
    s_instance->TableBordersAndShading();
}

TEST_F(ExDocumentBuilder, SetPreferredTypeConvertUtil)
{
    s_instance->SetPreferredTypeConvertUtil();
}

TEST_F(ExDocumentBuilder, InsertHyperlinkToLocalBookmark)
{
    s_instance->InsertHyperlinkToLocalBookmark();
}

TEST_F(ExDocumentBuilder, CursorPosition)
{
    s_instance->CursorPosition();
}

TEST_F(ExDocumentBuilder, MoveTo)
{
    s_instance->MoveTo();
}

TEST_F(ExDocumentBuilder, MoveToParagraph)
{
    s_instance->MoveToParagraph();
}

TEST_F(ExDocumentBuilder, MoveToCell)
{
    s_instance->MoveToCell();
}

TEST_F(ExDocumentBuilder, MoveToBookmark)
{
    s_instance->MoveToBookmark();
}

TEST_F(ExDocumentBuilder, BuildTable)
{
    s_instance->BuildTable();
}

TEST_F(ExDocumentBuilder, TableCellVerticalRotatedFarEastTextOrientation)
{
    s_instance->TableCellVerticalRotatedFarEastTextOrientation();
}

TEST_F(ExDocumentBuilder, InsertImageOriginalSize)
{
    s_instance->InsertImageOriginalSize();
}

TEST_F(ExDocumentBuilder, InsertTextInput)
{
    s_instance->InsertTextInput();
}

TEST_F(ExDocumentBuilder, InsertComboBox)
{
    s_instance->InsertComboBox();
}

TEST_F(ExDocumentBuilder, SignatureLineProviderId)
{
    s_instance->SignatureLineProviderId();
}

TEST_F(ExDocumentBuilder, SignatureLineInline)
{
    s_instance->SignatureLineInline();
}

TEST_F(ExDocumentBuilder, SetParagraphFormatting)
{
    s_instance->SetParagraphFormatting();
}

TEST_F(ExDocumentBuilder, SetCellFormatting)
{
    s_instance->SetCellFormatting();
}

TEST_F(ExDocumentBuilder, SetRowFormatting)
{
    s_instance->SetRowFormatting();
}

TEST_F(ExDocumentBuilder, InsertFootnote)
{
    s_instance->InsertFootnote();
}

TEST_F(ExDocumentBuilder, ApplyBordersAndShading)
{
    s_instance->ApplyBordersAndShading();
}

TEST_F(ExDocumentBuilder, DeleteRow)
{
    s_instance->DeleteRow();
}

using ExDocumentBuilder_AppendDocumentAndResolveStyles_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::AppendDocumentAndResolveStyles)>::type;

struct ExDocumentBuilder_AppendDocumentAndResolveStyles : public ExDocumentBuilder,
                                                          public ApiExamples::ExDocumentBuilder,
                                                          public ::testing::WithParamInterface<ExDocumentBuilder_AppendDocumentAndResolveStyles_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocumentBuilder_AppendDocumentAndResolveStyles, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AppendDocumentAndResolveStyles(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_AppendDocumentAndResolveStyles,
                         ::testing::ValuesIn(ExDocumentBuilder_AppendDocumentAndResolveStyles::TestCases()));

using ExDocumentBuilder_InsertDocumentAndResolveStyles_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::InsertDocumentAndResolveStyles)>::type;

struct ExDocumentBuilder_InsertDocumentAndResolveStyles : public ExDocumentBuilder,
                                                          public ApiExamples::ExDocumentBuilder,
                                                          public ::testing::WithParamInterface<ExDocumentBuilder_InsertDocumentAndResolveStyles_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocumentBuilder_InsertDocumentAndResolveStyles, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->InsertDocumentAndResolveStyles(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_InsertDocumentAndResolveStyles,
                         ::testing::ValuesIn(ExDocumentBuilder_InsertDocumentAndResolveStyles::TestCases()));

using ExDocumentBuilder_LoadDocumentWithListNumbering_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::LoadDocumentWithListNumbering)>::type;

struct ExDocumentBuilder_LoadDocumentWithListNumbering : public ExDocumentBuilder,
                                                         public ApiExamples::ExDocumentBuilder,
                                                         public ::testing::WithParamInterface<ExDocumentBuilder_LoadDocumentWithListNumbering_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocumentBuilder_LoadDocumentWithListNumbering, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->LoadDocumentWithListNumbering(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_LoadDocumentWithListNumbering, ::testing::ValuesIn(ExDocumentBuilder_LoadDocumentWithListNumbering::TestCases()));

using ExDocumentBuilder_IgnoreTextBoxes_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::IgnoreTextBoxes)>::type;

struct ExDocumentBuilder_IgnoreTextBoxes : public ExDocumentBuilder,
                                           public ApiExamples::ExDocumentBuilder,
                                           public ::testing::WithParamInterface<ExDocumentBuilder_IgnoreTextBoxes_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExDocumentBuilder_IgnoreTextBoxes, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreTextBoxes(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_IgnoreTextBoxes, ::testing::ValuesIn(ExDocumentBuilder_IgnoreTextBoxes::TestCases()));

using ExDocumentBuilder_MoveToField_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::MoveToField)>::type;

struct ExDocumentBuilder_MoveToField : public ExDocumentBuilder,
                                       public ApiExamples::ExDocumentBuilder,
                                       public ::testing::WithParamInterface<ExDocumentBuilder_MoveToField_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocumentBuilder_MoveToField, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MoveToField(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_MoveToField, ::testing::ValuesIn(ExDocumentBuilder_MoveToField::TestCases()));

TEST_F(ExDocumentBuilder, InsertOleObjectException)
{
    s_instance->InsertOleObjectException();
}

TEST_F(ExDocumentBuilder, InsertPieChart)
{
    s_instance->InsertPieChart();
}

TEST_F(ExDocumentBuilder, InsertChartRelativePosition)
{
    s_instance->InsertChartRelativePosition();
}

TEST_F(ExDocumentBuilder, InsertField)
{
    s_instance->InsertField();
}

using ExDocumentBuilder_InsertFieldAndUpdate_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::InsertFieldAndUpdate)>::type;

struct ExDocumentBuilder_InsertFieldAndUpdate : public ExDocumentBuilder,
                                                public ApiExamples::ExDocumentBuilder,
                                                public ::testing::WithParamInterface<ExDocumentBuilder_InsertFieldAndUpdate_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocumentBuilder_InsertFieldAndUpdate, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->InsertFieldAndUpdate(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_InsertFieldAndUpdate, ::testing::ValuesIn(ExDocumentBuilder_InsertFieldAndUpdate::TestCases()));

TEST_F(ExDocumentBuilder, FieldResultFormatting)
{
    s_instance->FieldResultFormatting_();
}

TEST_F(ExDocumentBuilder, InsertUnderline)
{
    s_instance->InsertUnderline();
}

TEST_F(ExDocumentBuilder, CurrentStory)
{
    s_instance->CurrentStory();
}

TEST_F(ExDocumentBuilder, InsertStyleSeparator)
{
    s_instance->InsertStyleSeparator();
}

TEST_F(ExDocumentBuilder, SmartStyleBehavior)
{
    s_instance->SmartStyleBehavior();
}

TEST_F(ExDocumentBuilder, EmphasesWarningSourceMarkdown)
{
    s_instance->EmphasesWarningSourceMarkdown();
}

TEST_F(ExDocumentBuilder, DoNotIgnoreHeaderFooter)
{
    s_instance->DoNotIgnoreHeaderFooter();
}

using ExDocumentBuilder_MarkdownDocumentTableContentAlignment_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::MarkdownDocumentTableContentAlignment)>::type;

struct ExDocumentBuilder_MarkdownDocumentTableContentAlignment
    : public ExDocumentBuilder,
      public ApiExamples::ExDocumentBuilder,
      public ::testing::WithParamInterface<ExDocumentBuilder_MarkdownDocumentTableContentAlignment_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(TableContentAlignment::Left),
            std::make_tuple(TableContentAlignment::Right),
            std::make_tuple(TableContentAlignment::Center),
            std::make_tuple(TableContentAlignment::Auto),
        };
    }
};

TEST_P(ExDocumentBuilder_MarkdownDocumentTableContentAlignment, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MarkdownDocumentTableContentAlignment(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocumentBuilder_MarkdownDocumentTableContentAlignment,
                         ::testing::ValuesIn(ExDocumentBuilder_MarkdownDocumentTableContentAlignment::TestCases()));

TEST_F(ExDocumentBuilder, RenameImages)
{
    s_instance->RenameImages();
}

TEST_F(ExDocumentBuilder, MarkdownDocumentTests)
{
    s_instance->MarkdownDocumentTests();
}

TEST_F(ExDocumentBuilder, InsertOnlineVideoCustomThumbnail)
{
    s_instance->InsertOnlineVideoCustomThumbnail();
}

TEST_F(ExDocumentBuilder, InsertOleObjectAsIcon)
{
    s_instance->InsertOleObjectAsIcon();
}

}} // namespace ApiExamples::gtest_test
