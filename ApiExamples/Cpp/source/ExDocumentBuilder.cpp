// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentBuilder.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/DateTimeFormatting/CalendarType.h>
#include <Aspose.Words.Cpp/Model/Fields/Format/GeneralFormat.h>
#include <system/collections/list.h>
#include <system/date_time.h>
#include <system/object.h>
#include <system/shared_ptr.h>
#include <system/string.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
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

using InsertHtmlWithFormatting_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::InsertHtmlWithFormatting)>::type;

struct InsertHtmlWithFormattingVP : public ExDocumentBuilder,
                                    public ApiExamples::ExDocumentBuilder,
                                    public ::testing::WithParamInterface<InsertHtmlWithFormatting_Args>
{
    static std::vector<InsertHtmlWithFormatting_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(InsertHtmlWithFormattingVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->InsertHtmlWithFormatting(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExDocumentBuilder, InsertHtmlWithFormattingVP, ::testing::ValuesIn(InsertHtmlWithFormattingVP::TestCases()));

TEST_F(ExDocumentBuilder, MathMl)
{
    s_instance->MathMl();
}

TEST_F(ExDocumentBuilder, InsertTextAndBookmark)
{
    s_instance->InsertTextAndBookmark();
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

TEST_F(ExDocumentBuilder, InsertFloatingImage)
{
    s_instance->InsertFloatingImage();
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

TEST_F(ExDocumentBuilder, AppendDocumentAndResolveStyles)
{
    s_instance->AppendDocumentAndResolveStyles();
}

using IgnoreTextBoxes_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::IgnoreTextBoxes)>::type;

struct IgnoreTextBoxesVP : public ExDocumentBuilder, public ApiExamples::ExDocumentBuilder, public ::testing::WithParamInterface<IgnoreTextBoxes_Args>
{
    static std::vector<IgnoreTextBoxes_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(IgnoreTextBoxesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreTextBoxes(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExDocumentBuilder, IgnoreTextBoxesVP, ::testing::ValuesIn(IgnoreTextBoxesVP::TestCases()));

using MoveToField_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::MoveToField)>::type;

struct MoveToFieldVP : public ExDocumentBuilder, public ApiExamples::ExDocumentBuilder, public ::testing::WithParamInterface<MoveToField_Args>
{
    static std::vector<MoveToField_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(MoveToFieldVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->MoveToField(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExDocumentBuilder, MoveToFieldVP, ::testing::ValuesIn(MoveToFieldVP::TestCases()));

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

using InsertFieldAndUpdate_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocumentBuilder::InsertFieldAndUpdate)>::type;

struct InsertFieldAndUpdateVP : public ExDocumentBuilder, public ApiExamples::ExDocumentBuilder, public ::testing::WithParamInterface<InsertFieldAndUpdate_Args>
{
    static std::vector<InsertFieldAndUpdate_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(InsertFieldAndUpdateVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->InsertFieldAndUpdate(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExDocumentBuilder, InsertFieldAndUpdateVP, ::testing::ValuesIn(InsertFieldAndUpdateVP::TestCases()));

TEST_F(ExDocumentBuilder, FieldResultFormatting)
{
    s_instance->FieldResultFormatting_();
}

TEST_F(ExDocumentBuilder, InsertVideoWithUrl)
{
    s_instance->InsertVideoWithUrl();
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

TEST_F(ExDocumentBuilder, DISABLED_InsertDocument)
{
    s_instance->InsertDocument();
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

TEST_F(ExDocumentBuilder, InsertOnlineVideo)
{
    s_instance->InsertOnlineVideo();
}

TEST_F(ExDocumentBuilder, InsertOnlineVideoCustomThumbnail)
{
    s_instance->InsertOnlineVideoCustomThumbnail();
}

}} // namespace ApiExamples::gtest_test
