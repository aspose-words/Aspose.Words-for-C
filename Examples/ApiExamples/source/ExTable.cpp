// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTable.h"

#include <testing/test_predicates.h>
#include <system/text/regularexpressions/regex.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/math.h>
#include <system/linq/enumerable.h>
#include <system/func.h>
#include <system/enumerator_adapter.h>
#include <system/enum_helpers.h>
#include <system/enum.h>
#include <system/details/dispose_guard.h>
#include <system/collections/ienumerator.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/rectangle.h>
#include <drawing/point.h>
#include <drawing/color.h>
#include <Aspose.Words.Cpp/Model/Text/TextOrientation.h>
#include <Aspose.Words.Cpp/Model/Text/TabStopCollection.h>
#include <Aspose.Words.Cpp/Model/Text/TabStop.h>
#include <Aspose.Words.Cpp/Model/Text/TabLeader.h>
#include <Aspose.Words.Cpp/Model/Text/TabAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/HeightRule.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Tables/TextWrapping.h>
#include <Aspose.Words.Cpp/Model/Tables/TableStyleOptions.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/TableAlignment.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidthType.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Model/Tables/CellVerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Tables/CellMerge.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/AutoFitBehavior.h>
#include <Aspose.Words.Cpp/Model/Styles/TableStyle.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/ConditionalStyleType.h>
#include <Aspose.Words.Cpp/Model/Styles/ConditionalStyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/ConditionalStyle.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilderOptions.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Borders/TextureIndex.h>
#include <Aspose.Words.Cpp/Model/Borders/Shading.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderType.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>

#include "TestUtil.h"
#include "DocumentHelper.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3808065182u, ::Aspose::Words::ApiExamples::ExTable, ThisTypeBaseTypesInfo);

int32_t ExTable::GetChildTableCount(System::SharedPtr<Aspose::Words::Tables::Table> table)
{
    int32_t childTableCount = 0;
    
    for (auto&& row : System::IterateOver<Aspose::Words::Tables::Row>(table->get_Rows()))
    {
        for (auto&& Cell : System::IterateOver<Aspose::Words::Tables::Cell>(row->get_Cells()))
        {
            System::SharedPtr<Aspose::Words::Tables::TableCollection> childTables = Cell->get_Tables();
            
            if (childTables->get_Count() > 0)
            {
                childTableCount++;
            }
        }
    }
    
    return childTableCount;
}

System::SharedPtr<Aspose::Words::Tables::Table> ExTable::CreateTable(System::SharedPtr<Aspose::Words::Document> doc, int32_t rowCount, int32_t cellCount, System::String cellText)
{
    auto table = System::MakeObject<Aspose::Words::Tables::Table>(doc);
    
    for (int32_t rowId = 1; rowId <= rowCount; rowId++)
    {
        auto row = System::MakeObject<Aspose::Words::Tables::Row>(doc);
        table->AppendChild<System::SharedPtr<Aspose::Words::Tables::Row>>(row);
        
        for (int32_t cellId = 1; cellId <= cellCount; cellId++)
        {
            auto cell = System::MakeObject<Aspose::Words::Tables::Cell>(doc);
            cell->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(System::MakeObject<Aspose::Words::Paragraph>(doc));
            cell->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, cellText));
            
            row->AppendChild<System::SharedPtr<Aspose::Words::Tables::Cell>>(cell);
        }
    }
    
    // You can use the "Title" and "Description" properties to add a title and description respectively to your table.
    // The table must have at least one row before we can use these properties.
    // These properties are meaningful for ISO / IEC 29500 compliant .docx documents (see the OoxmlCompliance class).
    // If we save the document to pre-ISO/IEC 29500 formats, Microsoft Word ignores these properties.
    table->set_Title(u"Aspose table title");
    table->set_Description(u"Aspose table description");
    
    return table;
}

void ExTable::TestCreateNestedTable(System::SharedPtr<Aspose::Words::Document> doc)
{
    System::SharedPtr<Aspose::Words::Tables::Table> outerTable = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    auto innerTable = System::ExplicitCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));
    
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
    ASSERT_EQ(1, outerTable->get_FirstRow()->get_FirstCell()->get_Tables()->get_Count());
    ASSERT_EQ(16, outerTable->GetChildNodes(Aspose::Words::NodeType::Cell, true)->get_Count());
    ASSERT_EQ(4, innerTable->GetChildNodes(Aspose::Words::NodeType::Cell, true)->get_Count());
    ASSERT_EQ(u"Aspose table title", innerTable->get_Title());
    ASSERT_EQ(u"Aspose table description", innerTable->get_Description());
}

System::String ExTable::PrintCellMergeType(System::SharedPtr<Aspose::Words::Tables::Cell> cell)
{
    bool isHorizontallyMerged = cell->get_CellFormat()->get_HorizontalMerge() != Aspose::Words::Tables::CellMerge::None;
    bool isVerticallyMerged = cell->get_CellFormat()->get_VerticalMerge() != Aspose::Words::Tables::CellMerge::None;
    System::String cellLocation = System::String::Format(u"R{0}, C{1}", cell->get_ParentRow()->get_ParentTable()->IndexOf(cell->get_ParentRow()) + 1, cell->get_ParentRow()->IndexOf(cell) + 1);
    
    if (isHorizontallyMerged && isVerticallyMerged)
    {
        return System::String::Format(u"The cell at {0} is both horizontally and vertically merged", cellLocation);
    }
    if (isHorizontallyMerged)
    {
        return System::String::Format(u"The cell at {0} is horizontally merged.", cellLocation);
    }
    
    return isVerticallyMerged ? System::String::Format(u"The cell at {0} is vertically merged", cellLocation) : System::String::Format(u"The cell at {0} is not merged", cellLocation);
}

void ExTable::MergeCells(System::SharedPtr<Aspose::Words::Tables::Cell> startCell, System::SharedPtr<Aspose::Words::Tables::Cell> endCell)
{
    System::SharedPtr<Aspose::Words::Tables::Table> parentTable = startCell->get_ParentRow()->get_ParentTable();
    
    // Find the row and cell indices for the start and end cells.
    System::Drawing::Point startCellPos(startCell->get_ParentRow()->IndexOf(startCell), parentTable->IndexOf(startCell->get_ParentRow()));
    System::Drawing::Point endCellPos(endCell->get_ParentRow()->IndexOf(endCell), parentTable->IndexOf(endCell->get_ParentRow()));
    
    // Create a range of cells to be merged based on these indices.
    // Inverse each index if the end cell is before the start cell.
    System::Drawing::Rectangle mergeRange(System::Math::Min(startCellPos.get_X(), endCellPos.get_X()), System::Math::Min(startCellPos.get_Y(), endCellPos.get_Y()), System::Math::Abs(endCellPos.get_X() - startCellPos.get_X()) + 1, System::Math::Abs(endCellPos.get_Y() - startCellPos.get_Y()) + 1);
    
    for (auto&& row : System::IterateOver<Aspose::Words::Tables::Row>(parentTable->get_Rows()))
    {
        for (auto&& cell : System::IterateOver<Aspose::Words::Tables::Cell>(row->get_Cells()))
        {
            System::Drawing::Point currentPos(row->IndexOf(cell), parentTable->IndexOf(row));
            
            // Check if the current cell is inside our merge range, then merge it.
            if (mergeRange.Contains(currentPos))
            {
                cell->get_CellFormat()->set_HorizontalMerge(currentPos.get_X() == mergeRange.get_X() ? Aspose::Words::Tables::CellMerge::First : Aspose::Words::Tables::CellMerge::Previous);
                cell->get_CellFormat()->set_VerticalMerge(currentPos.get_Y() == mergeRange.get_Y() ? Aspose::Words::Tables::CellMerge::First : Aspose::Words::Tables::CellMerge::Previous);
            }
        }
    }
}

void ExTable::ConvertTable(System::SharedPtr<Aspose::Words::Tables::Table> table)
{
    System::SharedPtr<Aspose::Words::Node> currentNode = table;
    for (auto&& row : System::IterateOver<Aspose::Words::Tables::Row>(table->get_Rows()))
    {
        for (auto&& cell : System::IterateOver<Aspose::Words::Tables::Cell>(row->get_Cells()))
        {
            // Get all nested tables within the current cell.
            System::SharedPtr<Aspose::Words::NodeCollection> nestedTables = cell->GetChildNodes(Aspose::Words::NodeType::Table, true);
            if (nestedTables->get_Count() != 0)
            {
                for (auto&& nestedTable : System::IterateOver<Aspose::Words::Tables::Table>(nestedTables))
                {
                    ConvertTable(nestedTable);
                }
            }
            
            // Get the text content of the cell and trim any whitespace.
            System::String cellText = cell->GetText().Trim();
            if (cellText == System::String::Empty)
            {
                break;
            }
            
            for (auto&& cellPara : System::IterateOver<Aspose::Words::Paragraph>(cell->get_Paragraphs()))
            {
                currentNode = table->get_ParentNode()->InsertAfter<System::SharedPtr<Aspose::Words::Node>>(System::ExplicitCast<Aspose::Words::Node>(cellPara)->Clone(true), currentNode);
            }
        }
    }
}

void ExTable::ConvertWith(System::String separator, System::SharedPtr<Aspose::Words::Tables::Table> table)
{
    auto doc = System::ExplicitCast<Aspose::Words::Document>(table->get_Document());
    System::SharedPtr<Aspose::Words::Node> currentPara = table->get_NextSibling();
    for (auto&& row : System::IterateOver<Aspose::Words::Tables::Row>(table->get_Rows()))
    {
        double tabStopWidth = 0;
        // By default MS Word adds 1.5 line spacing bitween paragraphs.
        (System::ExplicitCast<Aspose::Words::Paragraph>(currentPara))->get_ParagraphFormat()->set_LineSpacing(18);
        for (auto&& cell : System::IterateOver<Aspose::Words::Tables::Cell>(row->get_Cells()))
        {
            System::SharedPtr<Aspose::Words::NodeCollection> nestedTables = cell->GetChildNodes(Aspose::Words::NodeType::Table, true);
            // If there are nested tables, process each one.
            if (nestedTables->get_Count() != 0)
            {
                for (auto&& nestedTable : System::IterateOver<Aspose::Words::Tables::Table>(nestedTables))
                {
                    ConvertWith(separator, nestedTable);
                }
            }
            
            System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = cell->get_Paragraphs();
            for (auto&& paragraph : System::IterateOver<Aspose::Words::Paragraph>(paragraphs))
            {
                // If there's more than one paragraph and it's not the first, clone and insert it after the current paragraph.
                if (paragraphs->get_Count() > 1 && !System::ObjectExt::Equals(paragraph, cell->get_FirstParagraph()))
                {
                    System::SharedPtr<Aspose::Words::Node> node = currentPara->get_ParentNode()->InsertAfter<System::SharedPtr<Aspose::Words::Node>>(System::ExplicitCast<Aspose::Words::Node>(paragraph)->Clone(true), currentPara);
                    currentPara = node;
                }
                else if (currentPara->get_NodeType() == Aspose::Words::NodeType::Paragraph)
                {
                    // If the current cell is not the first cell, append a separator.
                    if (!cell->get_IsFirstCell())
                    {
                        (System::ExplicitCast<Aspose::Words::Paragraph>(currentPara))->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, separator));
                        // If the separator is a tab, calculate the tab stop position based on the width of the previous cell.
                        if (separator == Aspose::Words::ControlChar::Tab())
                        {
                            System::SharedPtr<Aspose::Words::Tables::Cell> previousCell = cell->get_PreviousCell();
                            if (previousCell != nullptr)
                            {
                                tabStopWidth += previousCell->get_CellFormat()->get_Width();
                            }
                            
                            // Add a tab stop at the calculated position.
                            auto tabStop = System::MakeObject<Aspose::Words::TabStop>(tabStopWidth, Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::None);
                            (System::ExplicitCast<Aspose::Words::Paragraph>(currentPara))->get_ParagraphFormat()->get_TabStops()->Add(tabStop);
                        }
                    }
                    
                    // Clone and append all child nodes of the paragraph to the current paragraph.
                    System::SharedPtr<Aspose::Words::NodeCollection> childNodes = paragraph->GetChildNodes(Aspose::Words::NodeType::Any, true);
                    if (childNodes->get_Count() > 0)
                    {
                        for (auto&& node : System::IterateOver(childNodes))
                        {
                            (System::ExplicitCast<Aspose::Words::Paragraph>(currentPara))->AppendChild<System::SharedPtr<Aspose::Words::Node>>(node->Clone(true));
                        }
                    }
                }
            }
        }
        
        currentPara = currentPara->get_ParentNode()->InsertAfter<System::SharedPtr<Aspose::Words::Paragraph>>(System::MakeObject<Aspose::Words::Paragraph>(doc), currentPara);
    }
}

int32_t ExTable::CalculateRowSpan(System::SharedPtr<Aspose::Words::Tables::Table> table, int32_t rowIndex, int32_t cellIndex)
{
    int32_t rowSpan = 1;
    for (int32_t i = rowIndex; i < table->get_Rows()->get_Count(); i++)
    {
        auto currentRow = table->get_Rows()->idx_get(i + 1);
        if (currentRow == nullptr)
        {
            break;
        }
        
        auto currentCell = currentRow->get_Cells()->idx_get(cellIndex);
        if (currentCell->get_CellFormat()->get_VerticalMerge() != Aspose::Words::Tables::CellMerge::Previous)
        {
            break;
        }
        
        rowSpan++;
    }
    return rowSpan;
}

System::SharedPtr<Aspose::Words::Tables::Cell> ExTable::CalculateColSpan(System::SharedPtr<Aspose::Words::Tables::Cell> cell, int32_t& colSpan)
{
    colSpan = 1;
    
    cell = cell->get_NextCell();
    while (cell != nullptr && cell->get_CellFormat()->get_HorizontalMerge() == Aspose::Words::Tables::CellMerge::Previous)
    {
        colSpan++;
        cell = cell->get_NextCell();
    }
    return cell;
}


namespace gtest_test
{

class ExTable : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExTable> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExTable>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExTable> ExTable::s_instance;

} // namespace gtest_test

void ExTable::CreateTable()
{
    //ExStart
    //ExFor:Table
    //ExFor:Row
    //ExFor:Cell
    //ExFor:Table.#ctor(DocumentBase)
    //ExSummary:Shows how to create a table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto table = System::MakeObject<Aspose::Words::Tables::Table>(doc);
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Tables::Table>>(table);
    
    // Tables contain rows, which contain cells, which may have paragraphs
    // with typical elements such as runs, shapes, and even other tables.
    // Calling the "EnsureMinimum" method on a table will ensure that
    // the table has at least one row, cell, and paragraph.
    auto firstRow = System::MakeObject<Aspose::Words::Tables::Row>(doc);
    table->AppendChild<System::SharedPtr<Aspose::Words::Tables::Row>>(firstRow);
    
    auto firstCell = System::MakeObject<Aspose::Words::Tables::Cell>(doc);
    firstRow->AppendChild<System::SharedPtr<Aspose::Words::Tables::Cell>>(firstCell);
    
    auto paragraph = System::MakeObject<Aspose::Words::Paragraph>(doc);
    firstCell->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(paragraph);
    
    // Add text to the first cell in the first row of the table.
    auto run = System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!");
    paragraph->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    doc->Save(get_ArtifactsDir() + u"Table.CreateTable.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.CreateTable.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(1, table->get_Rows()->get_Count());
    ASSERT_EQ(1, table->get_FirstRow()->get_Cells()->get_Count());
    ASSERT_EQ(u"Hello world!\a\a", table->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExTable, CreateTable)
{
    s_instance->CreateTable();
}

} // namespace gtest_test

void ExTable::Padding()
{
    //ExStart
    //ExFor:Table.LeftPadding
    //ExFor:Table.RightPadding
    //ExFor:Table.TopPadding
    //ExFor:Table.BottomPadding
    //ExSummary:Shows how to configure content padding in a table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Row 1, cell 1.");
    builder->InsertCell();
    builder->Write(u"Row 1, cell 2.");
    builder->EndTable();
    
    // For every cell in the table, set the distance between its contents and each of its borders.
    // This table will maintain the minimum padding distance by wrapping text.
    table->set_LeftPadding(30);
    table->set_RightPadding(60);
    table->set_TopPadding(10);
    table->set_BottomPadding(90);
    table->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPoints(250));
    
    doc->Save(get_ArtifactsDir() + u"DocumentBuilder.SetRowFormatting.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBuilder.SetRowFormatting.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASPOSE_ASSERT_EQ(30.0, table->get_LeftPadding());
    ASPOSE_ASSERT_EQ(60.0, table->get_RightPadding());
    ASPOSE_ASSERT_EQ(10.0, table->get_TopPadding());
    ASPOSE_ASSERT_EQ(90.0, table->get_BottomPadding());
}

namespace gtest_test
{

TEST_F(ExTable, Padding)
{
    s_instance->Padding();
}

} // namespace gtest_test

void ExTable::RowCellFormat()
{
    //ExStart
    //ExFor:Row.RowFormat
    //ExFor:RowFormat
    //ExFor:Cell.CellFormat
    //ExFor:CellFormat
    //ExFor:CellFormat.Shading
    //ExSummary:Shows how to modify the format of rows and cells in a table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"City");
    builder->InsertCell();
    builder->Write(u"Country");
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"London");
    builder->InsertCell();
    builder->Write(u"U.K.");
    builder->EndTable();
    
    // Use the first row's "RowFormat" property to modify the formatting
    // of the contents of all cells in this row.
    System::SharedPtr<Aspose::Words::Tables::RowFormat> rowFormat = table->get_FirstRow()->get_RowFormat();
    rowFormat->set_Height(25);
    rowFormat->get_Borders()->idx_get(Aspose::Words::BorderType::Bottom)->set_Color(System::Drawing::Color::get_Red());
    
    // Use the "CellFormat" property of the first cell in the last row to modify the formatting of that cell's contents.
    System::SharedPtr<Aspose::Words::Tables::CellFormat> cellFormat = table->get_LastRow()->get_FirstCell()->get_CellFormat();
    cellFormat->set_Width(100);
    cellFormat->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Orange());
    
    doc->Save(get_ArtifactsDir() + u"Table.RowCellFormat.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.RowCellFormat.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(u"City\aCountry\a\aLondon\aU.K.\a\a", table->GetText().Trim());
    
    rowFormat = table->get_FirstRow()->get_RowFormat();
    
    ASPOSE_ASSERT_EQ(25.0, rowFormat->get_Height());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), rowFormat->get_Borders()->idx_get(Aspose::Words::BorderType::Bottom)->get_Color().ToArgb());
    
    cellFormat = table->get_LastRow()->get_FirstCell()->get_CellFormat();
    
    ASPOSE_ASSERT_EQ(110.8, cellFormat->get_Width());
    ASSERT_EQ(System::Drawing::Color::get_Orange().ToArgb(), cellFormat->get_Shading()->get_BackgroundPatternColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExTable, RowCellFormat)
{
    s_instance->RowCellFormat();
}

} // namespace gtest_test

void ExTable::DisplayContentOfTables()
{
    //ExStart
    //ExFor:Cell
    //ExFor:CellCollection
    //ExFor:CellCollection.Item(Int32)
    //ExFor:CellCollection.ToArray
    //ExFor:Row
    //ExFor:Row.Cells
    //ExFor:RowCollection
    //ExFor:RowCollection.Item(Int32)
    //ExFor:RowCollection.ToArray
    //ExFor:Table
    //ExFor:Table.Rows
    //ExFor:TableCollection.Item(Int32)
    //ExFor:TableCollection.ToArray
    //ExSummary:Shows how to iterate through all tables in the document and print the contents of each cell.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    System::SharedPtr<Aspose::Words::Tables::TableCollection> tables = doc->get_FirstSection()->get_Body()->get_Tables();
    
    ASSERT_EQ(2, tables->ToArray()->get_Length());
    
    for (int32_t i = 0; i < tables->get_Count(); i++)
    {
        std::cout << System::String::Format(u"Start of Table {0}", i) << std::endl;
        
        System::SharedPtr<Aspose::Words::Tables::RowCollection> rows = tables->idx_get(i)->get_Rows();
        
        // We can use the "ToArray" method on a row collection to clone it into an array.
        ASPOSE_ASSERT_EQ(rows, rows->ToArray());
        ASPOSE_ASSERT_NS(rows, rows->ToArray());
        
        for (int32_t j = 0; j < rows->get_Count(); j++)
        {
            std::cout << System::String::Format(u"\tStart of Row {0}", j) << std::endl;
            
            System::SharedPtr<Aspose::Words::Tables::CellCollection> cells = rows->idx_get(j)->get_Cells();
            
            // We can use the "ToArray" method on a cell collection to clone it into an array.
            ASPOSE_ASSERT_EQ(cells, cells->ToArray());
            ASPOSE_ASSERT_NS(cells, cells->ToArray());
            
            for (int32_t k = 0; k < cells->get_Count(); k++)
            {
                System::String cellText = cells->idx_get(k)->ToString(Aspose::Words::SaveFormat::Text).Trim();
                std::cout << System::String::Format(u"\t\tContents of Cell:{0} = \"{1}\"", k, cellText) << std::endl;
            }
            
            std::cout << System::String::Format(u"\tEnd of Row {0}", j) << std::endl;
        }
        
        std::cout << System::String::Format(u"End of Table {0}\n", i) << std::endl;
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, DisplayContentOfTables)
{
    s_instance->DisplayContentOfTables();
}

} // namespace gtest_test

void ExTable::EnsureTableMinimum()
{
    //ExStart
    //ExFor:Table.EnsureMinimum
    //ExSummary:Shows how to ensure that a table node contains the nodes we need to add content.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto table = System::MakeObject<Aspose::Words::Tables::Table>(doc);
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Tables::Table>>(table);
    
    // Tables contain rows, which contain cells, which may contain paragraphs
    // with typical elements such as runs, shapes, and even other tables.
    // Our new table has none of these nodes, and we cannot add contents to it until it does.
    ASSERT_EQ(0, table->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
    
    // Calling the "EnsureMinimum" method on a table will ensure that
    // the table has at least one row and one cell with an empty paragraph.
    table->EnsureMinimum();
    table->get_FirstRow()->get_FirstCell()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!"));
    //ExEnd
    
    ASSERT_EQ(4, table->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
}

namespace gtest_test
{

TEST_F(ExTable, EnsureTableMinimum)
{
    s_instance->EnsureTableMinimum();
}

} // namespace gtest_test

void ExTable::EnsureRowMinimum()
{
    //ExStart
    //ExFor:Row.EnsureMinimum
    //ExSummary:Shows how to ensure a row node contains the nodes we need to begin adding content to it.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto table = System::MakeObject<Aspose::Words::Tables::Table>(doc);
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Tables::Table>>(table);
    auto row = System::MakeObject<Aspose::Words::Tables::Row>(doc);
    table->AppendChild<System::SharedPtr<Aspose::Words::Tables::Row>>(row);
    
    // Rows contain cells, containing paragraphs with typical elements such as runs, shapes, and even other tables.
    // Our new row has none of these nodes, and we cannot add contents to it until it does.
    ASSERT_EQ(0, row->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
    
    // Calling the "EnsureMinimum" method on a table will ensure that
    // the table has at least one cell with an empty paragraph.
    row->EnsureMinimum();
    row->get_FirstCell()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!"));
    //ExEnd
    
    ASSERT_EQ(3, row->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
}

namespace gtest_test
{

TEST_F(ExTable, EnsureRowMinimum)
{
    s_instance->EnsureRowMinimum();
}

} // namespace gtest_test

void ExTable::EnsureCellMinimum()
{
    //ExStart
    //ExFor:Cell.EnsureMinimum
    //ExSummary:Shows how to ensure a cell node contains the nodes we need to begin adding content to it.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto table = System::MakeObject<Aspose::Words::Tables::Table>(doc);
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Tables::Table>>(table);
    auto row = System::MakeObject<Aspose::Words::Tables::Row>(doc);
    table->AppendChild<System::SharedPtr<Aspose::Words::Tables::Row>>(row);
    auto cell = System::MakeObject<Aspose::Words::Tables::Cell>(doc);
    row->AppendChild<System::SharedPtr<Aspose::Words::Tables::Cell>>(cell);
    
    // Cells may contain paragraphs with typical elements such as runs, shapes, and even other tables.
    // Our new cell does not have any paragraphs, and we cannot add contents such as run and shape nodes to it until it does.
    ASSERT_EQ(0, cell->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
    
    // Calling the "EnsureMinimum" method on a cell will ensure that
    // the cell has at least one empty paragraph, which we can then add contents to.
    cell->EnsureMinimum();
    cell->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!"));
    //ExEnd
    
    ASSERT_EQ(2, cell->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
}

namespace gtest_test
{

TEST_F(ExTable, EnsureCellMinimum)
{
    s_instance->EnsureCellMinimum();
}

} // namespace gtest_test

void ExTable::SetOutlineBorders()
{
    //ExStart
    //ExFor:Table.Alignment
    //ExFor:TableAlignment
    //ExFor:Table.ClearBorders
    //ExFor:Table.ClearShading
    //ExFor:Table.SetBorder
    //ExFor:TextureIndex
    //ExFor:Table.SetShading
    //ExSummary:Shows how to apply an outline border to a table.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    // Align the table to the center of the page.
    table->set_Alignment(Aspose::Words::Tables::TableAlignment::Center);
    
    // Clear any existing borders and shading from the table.
    table->ClearBorders();
    table->ClearShading();
    
    // Add green borders to the outline of the table.
    table->SetBorder(Aspose::Words::BorderType::Left, Aspose::Words::LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
    table->SetBorder(Aspose::Words::BorderType::Right, Aspose::Words::LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
    table->SetBorder(Aspose::Words::BorderType::Top, Aspose::Words::LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
    table->SetBorder(Aspose::Words::BorderType::Bottom, Aspose::Words::LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
    
    // Fill the cells with a light green solid color.
    table->SetShading(Aspose::Words::TextureIndex::TextureSolid, System::Drawing::Color::get_LightGreen(), System::Drawing::Color::Empty);
    
    doc->Save(get_ArtifactsDir() + u"Table.SetOutlineBorders.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.SetOutlineBorders.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(Aspose::Words::Tables::TableAlignment::Center, table->get_Alignment());
    
    System::SharedPtr<Aspose::Words::BorderCollection> borders = table->get_FirstRow()->get_RowFormat()->get_Borders();
    
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), borders->get_Top()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), borders->get_Left()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), borders->get_Right()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), borders->get_Bottom()->get_Color().ToArgb());
    ASSERT_NE(System::Drawing::Color::get_Green().ToArgb(), borders->get_Horizontal()->get_Color().ToArgb());
    ASSERT_NE(System::Drawing::Color::get_Green().ToArgb(), borders->get_Vertical()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_LightGreen().ToArgb(), table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_ForegroundPatternColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExTable, SetOutlineBorders)
{
    s_instance->SetOutlineBorders();
}

} // namespace gtest_test

void ExTable::SetBorders()
{
    //ExStart
    //ExFor:Table.SetBorders
    //ExSummary:Shows how to format of all of a table's borders at once.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    // Clear all existing borders from the table.
    table->ClearBorders();
    
    // Set a single green line to serve as every outer and inner border of this table.
    table->SetBorders(Aspose::Words::LineStyle::Single, 1.5, System::Drawing::Color::get_Green());
    
    doc->Save(get_ArtifactsDir() + u"Table.SetBorders.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.SetBorders.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Top()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Left()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Right()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Bottom()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Horizontal()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Vertical()->get_Color().ToArgb());
}

namespace gtest_test
{

TEST_F(ExTable, SetBorders)
{
    s_instance->SetBorders();
}

} // namespace gtest_test

void ExTable::RowFormat()
{
    //ExStart
    //ExFor:RowFormat
    //ExFor:Row.RowFormat
    //ExSummary:Shows how to modify formatting of a table row.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    // Use the first row's "RowFormat" property to set formatting that modifies that entire row's appearance.
    System::SharedPtr<Aspose::Words::Tables::Row> firstRow = table->get_FirstRow();
    firstRow->get_RowFormat()->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::None);
    firstRow->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Auto);
    firstRow->get_RowFormat()->set_AllowBreakAcrossPages(true);
    
    doc->Save(get_ArtifactsDir() + u"Table.RowFormat.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.RowFormat.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(Aspose::Words::LineStyle::None, table->get_FirstRow()->get_RowFormat()->get_Borders()->get_LineStyle());
    ASSERT_EQ(Aspose::Words::HeightRule::Auto, table->get_FirstRow()->get_RowFormat()->get_HeightRule());
    ASSERT_TRUE(table->get_FirstRow()->get_RowFormat()->get_AllowBreakAcrossPages());
}

namespace gtest_test
{

TEST_F(ExTable, RowFormat)
{
    s_instance->RowFormat();
}

} // namespace gtest_test

void ExTable::CellFormat()
{
    //ExStart
    //ExFor:CellFormat
    //ExFor:Cell.CellFormat
    //ExSummary:Shows how to modify formatting of a table cell.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    System::SharedPtr<Aspose::Words::Tables::Cell> firstCell = table->get_FirstRow()->get_FirstCell();
    
    // Use a cell's "CellFormat" property to set formatting that modifies the appearance of that cell.
    firstCell->get_CellFormat()->set_Width(30);
    firstCell->get_CellFormat()->set_Orientation(Aspose::Words::TextOrientation::Downward);
    firstCell->get_CellFormat()->get_Shading()->set_ForegroundPatternColor(System::Drawing::Color::get_LightGreen());
    
    doc->Save(get_ArtifactsDir() + u"Table.CellFormat.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.CellFormat.docx");
    
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    ASPOSE_ASSERT_EQ(30, table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Width());
    ASSERT_EQ(Aspose::Words::TextOrientation::Downward, table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Orientation());
    ASSERT_EQ(System::Drawing::Color::get_LightGreen().ToArgb(), table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_ForegroundPatternColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExTable, CellFormat)
{
    s_instance->CellFormat();
}

} // namespace gtest_test

void ExTable::DistanceBetweenTableAndText()
{
    //ExStart
    //ExFor:Table.DistanceBottom
    //ExFor:Table.DistanceLeft
    //ExFor:Table.DistanceRight
    //ExFor:Table.DistanceTop
    //ExSummary:Shows how to set distance between table boundaries and text.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Table wrapped by text.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    ASPOSE_ASSERT_EQ(25.9, table->get_DistanceTop());
    ASPOSE_ASSERT_EQ(25.9, table->get_DistanceBottom());
    ASPOSE_ASSERT_EQ(17.3, table->get_DistanceLeft());
    ASPOSE_ASSERT_EQ(17.3, table->get_DistanceRight());
    
    // Set distance between table and surrounding text.
    table->set_DistanceLeft(24);
    table->set_DistanceRight(24);
    table->set_DistanceTop(3);
    table->set_DistanceBottom(3);
    
    doc->Save(get_ArtifactsDir() + u"Table.DistanceBetweenTableAndText.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, DistanceBetweenTableAndText)
{
    s_instance->DistanceBetweenTableAndText();
}

} // namespace gtest_test

void ExTable::Borders()
{
    //ExStart
    //ExFor:Table.ClearBorders
    //ExSummary:Shows how to remove all borders from a table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Hello world!");
    builder->EndTable();
    
    // Modify the color and thickness of the top border.
    System::SharedPtr<Aspose::Words::Border> topBorder = table->get_FirstRow()->get_RowFormat()->get_Borders()->idx_get(Aspose::Words::BorderType::Top);
    table->SetBorder(Aspose::Words::BorderType::Top, Aspose::Words::LineStyle::Double, 1.5, System::Drawing::Color::get_Red(), true);
    
    ASPOSE_ASSERT_EQ(1.5, topBorder->get_LineWidth());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), topBorder->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::LineStyle::Double, topBorder->get_LineStyle());
    
    // Clear the borders of all cells in the table, and then save the document.
    table->ClearBorders();
    doc->Save(get_ArtifactsDir() + u"Table.ClearBorders.docx");
    
    // Verify the values of the table's properties after re-opening the document.
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.ClearBorders.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    topBorder = table->get_FirstRow()->get_RowFormat()->get_Borders()->idx_get(Aspose::Words::BorderType::Top);
    
    ASPOSE_ASSERT_EQ(0.0, topBorder->get_LineWidth());
    ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), topBorder->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::LineStyle::None, topBorder->get_LineStyle());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, Borders)
{
    s_instance->Borders();
}

} // namespace gtest_test

void ExTable::ReplaceCellText()
{
    //ExStart
    //ExFor:Range.Replace(String, String, FindReplaceOptions)
    //ExSummary:Shows how to replace all instances of String of text in a table and cell.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Carrots");
    builder->InsertCell();
    builder->Write(u"50");
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"Potatoes");
    builder->InsertCell();
    builder->Write(u"50");
    builder->EndTable();
    
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    options->set_MatchCase(true);
    options->set_FindWholeWordsOnly(true);
    
    // Perform a find-and-replace operation on an entire table.
    table->get_Range()->Replace(u"Carrots", u"Eggs", options);
    
    // Perform a find-and-replace operation on the last cell of the last row of the table.
    table->get_LastRow()->get_LastCell()->get_Range()->Replace(u"50", u"20", options);
    
    ASSERT_EQ(System::String(u"Eggs\a50\a\a") + u"Potatoes\a20\a\a", table->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, ReplaceCellText)
{
    s_instance->ReplaceCellText();
}

} // namespace gtest_test

void ExTable::RemoveParagraphTextAndMark(bool isSmartParagraphBreakReplacement)
{
    //ExStart
    //ExFor:FindReplaceOptions.SmartParagraphBreakReplacement
    //ExSummary:Shows how to remove paragraph from a table cell with a nested table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create table with paragraph and inner table in first cell.
    builder->StartTable();
    builder->InsertCell();
    builder->Write(u"TEXT1");
    builder->StartTable();
    builder->InsertCell();
    builder->EndTable();
    builder->EndTable();
    builder->Writeln();
    
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    // When the following option is set to 'true', Aspose.Words will remove paragraph's text
    // completely with its paragraph mark. Otherwise, Aspose.Words will mimic Word and remove
    // only paragraph's text and leaves the paragraph mark intact (when a table follows the text).
    options->set_SmartParagraphBreakReplacement(isSmartParagraphBreakReplacement);
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"TEXT1&p"), u"", options);
    
    doc->Save(get_ArtifactsDir() + u"Table.RemoveParagraphTextAndMark.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.RemoveParagraphTextAndMark.docx");
    
    ASSERT_EQ(isSmartParagraphBreakReplacement ? 1 : 2, doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->get_Paragraphs()->get_Count());
}

namespace gtest_test
{

using ExTable_RemoveParagraphTextAndMark_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExTable::RemoveParagraphTextAndMark)>::type;

struct ExTable_RemoveParagraphTextAndMark : public ExTable, public Aspose::Words::ApiExamples::ExTable, public ::testing::WithParamInterface<ExTable_RemoveParagraphTextAndMark_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExTable_RemoveParagraphTextAndMark, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RemoveParagraphTextAndMark(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTable_RemoveParagraphTextAndMark, ::testing::ValuesIn(ExTable_RemoveParagraphTextAndMark::TestCases()));

} // namespace gtest_test

void ExTable::PrintTableRange()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    // The range text will include control characters such as "\a" for a cell.
    // You can call ToString on the desired node to retrieve the plain text content.
    
    // Print the plain text range of the table to the screen.
    std::cout << "Contents of the table: " << std::endl;
    std::cout << table->get_Range()->get_Text() << std::endl;
    
    // Print the contents of the second row to the screen.
    std::cout << "\nContents of the row: " << std::endl;
    std::cout << table->get_Rows()->idx_get(1)->get_Range()->get_Text() << std::endl;
    
    // Print the contents of the last cell in the table to the screen.
    std::cout << "\nContents of the cell: " << std::endl;
    std::cout << table->get_LastRow()->get_LastCell()->get_Range()->get_Text() << std::endl;
    
    ASSERT_EQ(u"\aColumn 1\aColumn 2\aColumn 3\aColumn 4\a\a", table->get_Rows()->idx_get(1)->get_Range()->get_Text());
    ASSERT_EQ(u"Cell 12 contents\a", table->get_LastRow()->get_LastCell()->get_Range()->get_Text());
}

namespace gtest_test
{

TEST_F(ExTable, PrintTableRange)
{
    s_instance->PrintTableRange();
}

} // namespace gtest_test

void ExTable::CloneTable()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    auto tableClone = System::ExplicitCast<Aspose::Words::Tables::Table>(System::ExplicitCast<Aspose::Words::Node>(table)->Clone(true));
    
    // Insert the cloned table into the document after the original.
    table->get_ParentNode()->InsertAfter<System::SharedPtr<Aspose::Words::Tables::Table>>(tableClone, table);
    
    // Insert an empty paragraph between the two tables.
    table->get_ParentNode()->InsertAfter<System::SharedPtr<Aspose::Words::Paragraph>>(System::MakeObject<Aspose::Words::Paragraph>(doc), table);
    
    doc->Save(get_ArtifactsDir() + u"Table.CloneTable.doc");
    
    ASSERT_EQ(3, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
    ASSERT_EQ(table->get_Range()->get_Text(), tableClone->get_Range()->get_Text());
    
    for (auto&& cell : System::IterateOver<Aspose::Words::Tables::Cell>(tableClone->GetChildNodes(Aspose::Words::NodeType::Cell, true)))
    {
        cell->RemoveAllChildren();
    }
    
    ASSERT_EQ(System::String::Empty, tableClone->ToString(Aspose::Words::SaveFormat::Text).Trim());
}

namespace gtest_test
{

TEST_F(ExTable, CloneTable)
{
    s_instance->CloneTable();
}

} // namespace gtest_test

void ExTable::AllowBreakAcrossPages(bool allowBreakAcrossPages)
{
    //ExStart
    //ExFor:RowFormat.AllowBreakAcrossPages
    //ExSummary:Shows how to disable rows breaking across pages for every row in a table.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Table spanning two pages.docx");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    // Set the "AllowBreakAcrossPages" property to "false" to keep the row
    // in one piece if a table spans two pages, which break up along that row.
    // If the row is too big to fit in one page, Microsoft Word will push it down to the next page.
    // Set the "AllowBreakAcrossPages" property to "true" to allow the row to break up across two pages.
    for (auto&& row : System::IterateOver<Aspose::Words::Tables::Row>(table))
    {
        row->get_RowFormat()->set_AllowBreakAcrossPages(allowBreakAcrossPages);
    }
    
    doc->Save(get_ArtifactsDir() + u"Table.AllowBreakAcrossPages.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.AllowBreakAcrossPages.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(3, table->get_Rows()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> r)>>([&allowBreakAcrossPages](System::SharedPtr<Aspose::Words::Node> r) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Tables::Row>(r))->get_RowFormat()->get_AllowBreakAcrossPages() == allowBreakAcrossPages;
    }))));
}

namespace gtest_test
{

using ExTable_AllowBreakAcrossPages_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExTable::AllowBreakAcrossPages)>::type;

struct ExTable_AllowBreakAcrossPages : public ExTable, public Aspose::Words::ApiExamples::ExTable, public ::testing::WithParamInterface<ExTable_AllowBreakAcrossPages_Args>
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

TEST_P(ExTable_AllowBreakAcrossPages, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AllowBreakAcrossPages(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTable_AllowBreakAcrossPages, ::testing::ValuesIn(ExTable_AllowBreakAcrossPages::TestCases()));

} // namespace gtest_test

void ExTable::AllowAutoFitOnTable(bool allowAutoFit)
{
    //ExStart
    //ExFor:Table.AllowAutoFit
    //ExSummary:Shows how to enable/disable automatic table cell resizing.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->get_CellFormat()->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPoints(100));
    builder->Write(System::String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") + u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    
    builder->InsertCell();
    builder->get_CellFormat()->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::Auto());
    builder->Write(System::String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") + u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    builder->EndRow();
    builder->EndTable();
    
    // Set the "AllowAutoFit" property to "false" to get the table to maintain the dimensions
    // of all its rows and cells, and truncate contents if they get too large to fit.
    // Set the "AllowAutoFit" property to "true" to allow the table to change its cells' width and height
    // to accommodate their contents.
    table->set_AllowAutoFit(allowAutoFit);
    
    doc->Save(get_ArtifactsDir() + u"Table.AllowAutoFitOnTable.html");
    //ExEnd
    
    if (allowAutoFit)
    {
        Aspose::Words::ApiExamples::TestUtil::FileContainsString(u"<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-right:0.5pt single\">", get_ArtifactsDir() + u"Table.AllowAutoFitOnTable.html");
        Aspose::Words::ApiExamples::TestUtil::FileContainsString(u"<td style=\"border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-left:0.5pt single\">", get_ArtifactsDir() + u"Table.AllowAutoFitOnTable.html");
    }
    else
    {
        Aspose::Words::ApiExamples::TestUtil::FileContainsString(u"<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-right:0.5pt single\">", get_ArtifactsDir() + u"Table.AllowAutoFitOnTable.html");
        Aspose::Words::ApiExamples::TestUtil::FileContainsString(u"<td style=\"width:7.2pt; border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-left:0.5pt single\">", get_ArtifactsDir() + u"Table.AllowAutoFitOnTable.html");
    }
}

namespace gtest_test
{

using ExTable_AllowAutoFitOnTable_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExTable::AllowAutoFitOnTable)>::type;

struct ExTable_AllowAutoFitOnTable : public ExTable, public Aspose::Words::ApiExamples::ExTable, public ::testing::WithParamInterface<ExTable_AllowAutoFitOnTable_Args>
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

TEST_P(ExTable_AllowAutoFitOnTable, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AllowAutoFitOnTable(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTable_AllowAutoFitOnTable, ::testing::ValuesIn(ExTable_AllowAutoFitOnTable::TestCases()));

} // namespace gtest_test

void ExTable::KeepTableTogether()
{
    //ExStart
    //ExFor:ParagraphFormat.KeepWithNext
    //ExFor:Row.IsLastRow
    //ExFor:Paragraph.IsEndOfCell
    //ExFor:Paragraph.IsInCell
    //ExFor:Cell.ParentRow
    //ExFor:Cell.Paragraphs
    //ExSummary:Shows how to set a table to stay together on the same page.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Table spanning two pages.docx");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    // Enabling KeepWithNext for every paragraph in the table except for the
    // last ones in the last row will prevent the table from splitting across multiple pages.
    for (auto&& cell : System::IterateOver<Aspose::Words::Tables::Cell>(table->GetChildNodes(Aspose::Words::NodeType::Cell, true)))
    {
        for (auto&& para : System::IterateOver<Aspose::Words::Paragraph>(cell->get_Paragraphs()))
        {
            ASSERT_TRUE(para->get_IsInCell());
            
            if (!(cell->get_ParentRow()->get_IsLastRow() && para->get_IsEndOfCell()))
            {
                para->get_ParagraphFormat()->set_KeepWithNext(true);
            }
        }
    }
    
    doc->Save(get_ArtifactsDir() + u"Table.KeepTableTogether.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.KeepTableTogether.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    for (auto&& para : System::IterateOver<Aspose::Words::Paragraph>(table->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)))
    {
        if (para->get_IsEndOfCell() && (System::ExplicitCast<Aspose::Words::Tables::Cell>(para->get_ParentNode()))->get_ParentRow()->get_IsLastRow())
        {
            ASSERT_FALSE(para->get_ParagraphFormat()->get_KeepWithNext());
        }
        else
        {
            ASSERT_TRUE(para->get_ParagraphFormat()->get_KeepWithNext());
        }
    }
}

namespace gtest_test
{

TEST_F(ExTable, KeepTableTogether)
{
    s_instance->KeepTableTogether();
}

} // namespace gtest_test

void ExTable::GetIndexOfTableElements()
{
    //ExStart
    //ExFor:NodeCollection.IndexOf(Node)
    //ExSummary:Shows how to get the index of a node in a collection.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    System::SharedPtr<Aspose::Words::NodeCollection> allTables = doc->GetChildNodes(Aspose::Words::NodeType::Table, true);
    
    ASSERT_EQ(0, allTables->IndexOf(table));
    
    System::SharedPtr<Aspose::Words::Tables::Row> row = table->get_Rows()->idx_get(2);
    
    ASSERT_EQ(2, table->IndexOf(row));
    
    System::SharedPtr<Aspose::Words::Tables::Cell> cell = row->get_LastCell();
    
    ASSERT_EQ(4, row->IndexOf(cell));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, GetIndexOfTableElements)
{
    s_instance->GetIndexOfTableElements();
}

} // namespace gtest_test

void ExTable::GetPreferredWidthTypeAndValue()
{
    //ExStart
    //ExFor:PreferredWidthType
    //ExFor:PreferredWidth.Type
    //ExFor:PreferredWidth.Value
    //ExSummary:Shows how to verify the preferred width type and value of a table cell.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    System::SharedPtr<Aspose::Words::Tables::Cell> firstCell = table->get_FirstRow()->get_FirstCell();
    
    ASSERT_EQ(Aspose::Words::Tables::PreferredWidthType::Percent, firstCell->get_CellFormat()->get_PreferredWidth()->get_Type());
    ASPOSE_ASSERT_EQ(11.16, firstCell->get_CellFormat()->get_PreferredWidth()->get_Value());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, GetPreferredWidthTypeAndValue)
{
    s_instance->GetPreferredWidthTypeAndValue();
}

} // namespace gtest_test

void ExTable::AllowCellSpacing(bool allowCellSpacing)
{
    //ExStart
    //ExFor:Table.AllowCellSpacing
    //ExFor:Table.CellSpacing
    //ExSummary:Shows how to enable spacing between individual cells in a table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Animal");
    builder->InsertCell();
    builder->Write(u"Class");
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"Dog");
    builder->InsertCell();
    builder->Write(u"Mammal");
    builder->EndTable();
    
    table->set_CellSpacing(3);
    
    // Set the "AllowCellSpacing" property to "true" to enable spacing between cells
    // with a magnitude equal to the value of the "CellSpacing" property, in points.
    // Set the "AllowCellSpacing" property to "false" to disable cell spacing
    // and ignore the value of the "CellSpacing" property.
    table->set_AllowCellSpacing(allowCellSpacing);
    
    doc->Save(get_ArtifactsDir() + u"Table.AllowCellSpacing.html");
    
    // Adjusting the "CellSpacing" property will automatically enable cell spacing.
    table->set_CellSpacing(5);
    
    ASSERT_TRUE(table->get_AllowCellSpacing());
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.AllowCellSpacing.html");
    table = System::ExplicitCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    
    ASPOSE_ASSERT_EQ(allowCellSpacing, table->get_AllowCellSpacing());
    
    if (allowCellSpacing)
    {
        ASPOSE_ASSERT_EQ(3.0, table->get_CellSpacing());
    }
    else
    {
        ASPOSE_ASSERT_EQ(0.0, table->get_CellSpacing());
    }
    
    Aspose::Words::ApiExamples::TestUtil::FileContainsString(allowCellSpacing ? u"<td style=\"border-style:solid; border-width:0.75pt; padding-right:5.4pt; padding-left:5.4pt; vertical-align:top; -aw-border:0.5pt single\">" : System::String(u"<td style=\"border-right-style:solid; border-right-width:0.75pt; border-bottom-style:solid; border-bottom-width:0.75pt; ") + u"padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-bottom:0.5pt single; -aw-border-right:0.5pt single\">", get_ArtifactsDir() + u"Table.AllowCellSpacing.html");
}

namespace gtest_test
{

using ExTable_AllowCellSpacing_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExTable::AllowCellSpacing)>::type;

struct ExTable_AllowCellSpacing : public ExTable, public Aspose::Words::ApiExamples::ExTable, public ::testing::WithParamInterface<ExTable_AllowCellSpacing_Args>
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

TEST_P(ExTable_AllowCellSpacing, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AllowCellSpacing(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTable_AllowCellSpacing, ::testing::ValuesIn(ExTable_AllowCellSpacing::TestCases()));

} // namespace gtest_test

void ExTable::CreateNestedTable()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create the outer table with three rows and four columns, and then add it to the document.
    System::SharedPtr<Aspose::Words::Tables::Table> outerTable = CreateTable(doc, 3, 4, u"Outer Table");
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Tables::Table>>(outerTable);
    
    // Create another table with two rows and two columns and then insert it into the first table's first cell.
    System::SharedPtr<Aspose::Words::Tables::Table> innerTable = CreateTable(doc, 2, 2, u"Inner Table");
    outerTable->get_FirstRow()->get_FirstCell()->AppendChild<System::SharedPtr<Aspose::Words::Tables::Table>>(innerTable);
    
    doc->Save(get_ArtifactsDir() + u"Table.CreateNestedTable.docx");
    TestCreateNestedTable(System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.CreateNestedTable.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExTable, CreateNestedTable)
{
    s_instance->CreateNestedTable();
}

} // namespace gtest_test

void ExTable::CheckCellsMerged()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Table with merged cells.docx");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    for (auto&& row : System::IterateOver<Aspose::Words::Tables::Row>(table->get_Rows()))
    {
        for (auto&& cell : System::IterateOver<Aspose::Words::Tables::Cell>(row->get_Cells()))
        {
            std::cout << PrintCellMergeType(cell) << std::endl;
        }
    }
    ASSERT_EQ(u"The cell at R1, C1 is vertically merged", PrintCellMergeType(table->get_FirstRow()->get_FirstCell()));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExTable, CheckCellsMerged)
{
    s_instance->CheckCellsMerged();
}

} // namespace gtest_test

void ExTable::MergeCellRange()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    // We want to merge the range of cells found in between these two cells.
    System::SharedPtr<Aspose::Words::Tables::Cell> cellStartRange = table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2);
    System::SharedPtr<Aspose::Words::Tables::Cell> cellEndRange = table->get_Rows()->idx_get(3)->get_Cells()->idx_get(3);
    
    // Merge all the cells between the two specified cells into one.
    MergeCells(cellStartRange, cellEndRange);
    
    doc->Save(get_ArtifactsDir() + u"Table.MergeCellRange.doc");
    
    int32_t mergedCellsCount = 0;
    for (auto&& node : System::IterateOver(table->GetChildNodes(Aspose::Words::NodeType::Cell, true)))
    {
        auto cell = System::ExplicitCast<Aspose::Words::Tables::Cell>(node);
        if (cell->get_CellFormat()->get_HorizontalMerge() != Aspose::Words::Tables::CellMerge::None || cell->get_CellFormat()->get_VerticalMerge() != Aspose::Words::Tables::CellMerge::None)
        {
            mergedCellsCount++;
        }
    }
    
    ASSERT_EQ(4, mergedCellsCount);
    ASSERT_TRUE(table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2)->get_CellFormat()->get_HorizontalMerge() == Aspose::Words::Tables::CellMerge::First);
    ASSERT_TRUE(table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2)->get_CellFormat()->get_VerticalMerge() == Aspose::Words::Tables::CellMerge::First);
    ASSERT_TRUE(table->get_Rows()->idx_get(3)->get_Cells()->idx_get(3)->get_CellFormat()->get_HorizontalMerge() == Aspose::Words::Tables::CellMerge::Previous);
    ASSERT_TRUE(table->get_Rows()->idx_get(3)->get_Cells()->idx_get(3)->get_CellFormat()->get_VerticalMerge() == Aspose::Words::Tables::CellMerge::Previous);
}

namespace gtest_test
{

TEST_F(ExTable, MergeCellRange)
{
    s_instance->MergeCellRange();
}

} // namespace gtest_test

void ExTable::CombineTables()
{
    //ExStart
    //ExFor:Cell.CellFormat
    //ExFor:CellFormat.Borders
    //ExFor:Table.Rows
    //ExFor:Table.FirstRow
    //ExFor:CellFormat.ClearFormatting
    //ExFor:CompositeNode.HasChildNodes
    //ExSummary:Shows how to combine the rows from two tables into one.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    
    // Below are two ways of getting a table from a document.
    // 1 -  From the "Tables" collection of a Body node:
    System::SharedPtr<Aspose::Words::Tables::Table> firstTable = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    // 2 -  Using the "GetChild" method:
    auto secondTable = System::ExplicitCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));
    
    // Append all rows from the current table to the next.
    while (secondTable->get_HasChildNodes())
    {
        firstTable->get_Rows()->Add(secondTable->get_FirstRow());
    }
    
    // Remove the empty table container.
    secondTable->Remove();
    
    doc->Save(get_ArtifactsDir() + u"Table.CombineTables.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.CombineTables.docx");
    
    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
    ASSERT_EQ(9, doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_Rows()->get_Count());
    ASSERT_EQ(42, doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->GetChildNodes(Aspose::Words::NodeType::Cell, true)->get_Count());
}

namespace gtest_test
{

TEST_F(ExTable, CombineTables)
{
    s_instance->CombineTables();
}

} // namespace gtest_test

void ExTable::SplitTable()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Table> firstTable = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    // We will split the table at the third row (inclusive).
    System::SharedPtr<Aspose::Words::Tables::Row> row = firstTable->get_Rows()->idx_get(2);
    
    // Create a new container for the split table.
    auto table = System::ExplicitCast<Aspose::Words::Tables::Table>(System::ExplicitCast<Aspose::Words::Node>(firstTable)->Clone(false));
    
    // Insert the container after the original.
    firstTable->get_ParentNode()->InsertAfter<System::SharedPtr<Aspose::Words::Tables::Table>>(table, firstTable);
    
    // Add a buffer paragraph to ensure the tables stay apart.
    firstTable->get_ParentNode()->InsertAfter<System::SharedPtr<Aspose::Words::Paragraph>>(System::MakeObject<Aspose::Words::Paragraph>(doc), firstTable);
    
    System::SharedPtr<Aspose::Words::Tables::Row> currentRow;
    
    do
    {
        currentRow = firstTable->get_LastRow();
        table->PrependChild<System::SharedPtr<Aspose::Words::Tables::Row>>(currentRow);
    } while (currentRow != row);
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    
    ASPOSE_ASSERT_EQ(row, table->get_FirstRow());
    ASSERT_EQ(2, firstTable->get_Rows()->get_Count());
    ASSERT_EQ(3, table->get_Rows()->get_Count());
    ASSERT_EQ(3, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
}

namespace gtest_test
{

TEST_F(ExTable, SplitTable)
{
    s_instance->SplitTable();
}

} // namespace gtest_test

void ExTable::WrapText()
{
    //ExStart
    //ExFor:Table.TextWrapping
    //ExFor:TextWrapping
    //ExSummary:Shows how to work with table text wrapping.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Cell 1");
    builder->InsertCell();
    builder->Write(u"Cell 2");
    builder->EndTable();
    table->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPoints(300));
    
    builder->get_Font()->set_Size(16);
    builder->Writeln(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    
    // Set the "TextWrapping" property to "TextWrapping.Around" to get the table to wrap text around it,
    // and push it down into the paragraph below by setting the position.
    table->set_TextWrapping(Aspose::Words::Tables::TextWrapping::Around);
    table->set_AbsoluteHorizontalDistance(100);
    table->set_AbsoluteVerticalDistance(20);
    
    doc->Save(get_ArtifactsDir() + u"Table.WrapText.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.WrapText.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(Aspose::Words::Tables::TextWrapping::Around, table->get_TextWrapping());
    ASPOSE_ASSERT_EQ(100.0, table->get_AbsoluteHorizontalDistance());
    ASPOSE_ASSERT_EQ(20.0, table->get_AbsoluteVerticalDistance());
}

namespace gtest_test
{

TEST_F(ExTable, WrapText)
{
    s_instance->WrapText();
}

} // namespace gtest_test

void ExTable::GetFloatingTableProperties()
{
    //ExStart
    //ExFor:Table.HorizontalAnchor
    //ExFor:Table.VerticalAnchor
    //ExFor:Table.AllowOverlap
    //ExFor:ShapeBase.AllowOverlap
    //ExSummary:Shows how to work with floating tables properties.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Table wrapped by text.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    if (table->get_TextWrapping() == Aspose::Words::Tables::TextWrapping::Around)
    {
        ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, table->get_HorizontalAnchor());
        ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, table->get_VerticalAnchor());
        ASPOSE_ASSERT_EQ(false, table->get_AllowOverlap());
        
        // Only Margin, Page, Column available in RelativeHorizontalPosition for HorizontalAnchor setter.
        // The ArgumentException will be thrown for any other values.
        table->set_HorizontalAnchor(Aspose::Words::Drawing::RelativeHorizontalPosition::Column);
        
        // Only Margin, Page, Paragraph available in RelativeVerticalPosition for VerticalAnchor setter.
        // The ArgumentException will be thrown for any other values.
        table->set_VerticalAnchor(Aspose::Words::Drawing::RelativeVerticalPosition::Page);
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, GetFloatingTableProperties)
{
    s_instance->GetFloatingTableProperties();
}

} // namespace gtest_test

void ExTable::ChangeFloatingTableProperties()
{
    //ExStart
    //ExFor:Table.RelativeHorizontalAlignment
    //ExFor:Table.RelativeVerticalAlignment
    //ExFor:Table.AbsoluteHorizontalDistance
    //ExFor:Table.AbsoluteVerticalDistance
    //ExSummary:Shows how set the location of floating tables.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Table 1, cell 1");
    builder->EndTable();
    table->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPoints(300));
    
    // Set the table's location to a place on the page, such as, in this case, the bottom right corner.
    table->set_RelativeVerticalAlignment(Aspose::Words::Drawing::VerticalAlignment::Bottom);
    table->set_RelativeHorizontalAlignment(Aspose::Words::Drawing::HorizontalAlignment::Right);
    
    table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Table 2, cell 1");
    builder->EndTable();
    table->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPoints(300));
    
    // We can also set a horizontal and vertical offset in points from the paragraph's location where we inserted the table.
    table->set_AbsoluteVerticalDistance(50);
    table->set_AbsoluteHorizontalDistance(100);
    
    doc->Save(get_ArtifactsDir() + u"Table.ChangeFloatingTableProperties.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.ChangeFloatingTableProperties.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(Aspose::Words::Drawing::VerticalAlignment::Bottom, table->get_RelativeVerticalAlignment());
    ASSERT_EQ(Aspose::Words::Drawing::HorizontalAlignment::Right, table->get_RelativeHorizontalAlignment());
    
    table = System::ExplicitCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));
    
    ASPOSE_ASSERT_EQ(50.0, table->get_AbsoluteVerticalDistance());
    ASPOSE_ASSERT_EQ(100.0, table->get_AbsoluteHorizontalDistance());
}

namespace gtest_test
{

TEST_F(ExTable, ChangeFloatingTableProperties)
{
    s_instance->ChangeFloatingTableProperties();
}

} // namespace gtest_test

void ExTable::TableStyleCreation()
{
    //ExStart
    //ExFor:Table.Bidi
    //ExFor:Table.CellSpacing
    //ExFor:Table.Style
    //ExFor:Table.StyleName
    //ExFor:TableStyle
    //ExFor:TableStyle.AllowBreakAcrossPages
    //ExFor:TableStyle.Bidi
    //ExFor:TableStyle.CellSpacing
    //ExFor:TableStyle.BottomPadding
    //ExFor:TableStyle.LeftPadding
    //ExFor:TableStyle.RightPadding
    //ExFor:TableStyle.TopPadding
    //ExFor:TableStyle.Shading
    //ExFor:TableStyle.Borders
    //ExFor:TableStyle.VerticalAlignment
    //ExSummary:Shows how to create custom style settings for the table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Name");
    builder->InsertCell();
    builder->Write(u"مرحبًا");
    builder->EndRow();
    builder->InsertCell();
    builder->InsertCell();
    builder->EndTable();
    
    auto tableStyle = System::ExplicitCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle1"));
    tableStyle->set_AllowBreakAcrossPages(true);
    tableStyle->set_Bidi(true);
    tableStyle->set_CellSpacing(5);
    tableStyle->set_BottomPadding(20);
    tableStyle->set_LeftPadding(5);
    tableStyle->set_RightPadding(10);
    tableStyle->set_TopPadding(20);
    tableStyle->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_AntiqueWhite());
    tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Blue());
    tableStyle->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::DotDash);
    tableStyle->set_VerticalAlignment(Aspose::Words::Tables::CellVerticalAlignment::Center);
    
    table->set_Style(tableStyle);
    
    // Setting the style properties of a table may affect the properties of the table itself.
    ASSERT_TRUE(table->get_Bidi());
    ASPOSE_ASSERT_EQ(5.0, table->get_CellSpacing());
    ASSERT_EQ(u"MyTableStyle1", table->get_StyleName());
    
    doc->Save(get_ArtifactsDir() + u"Table.TableStyleCreation.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.TableStyleCreation.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_TRUE(table->get_Bidi());
    ASPOSE_ASSERT_EQ(5.0, table->get_CellSpacing());
    ASSERT_EQ(u"MyTableStyle1", table->get_StyleName());
    ASPOSE_ASSERT_EQ(20.0, tableStyle->get_BottomPadding());
    ASPOSE_ASSERT_EQ(5.0, tableStyle->get_LeftPadding());
    ASPOSE_ASSERT_EQ(10.0, tableStyle->get_RightPadding());
    ASPOSE_ASSERT_EQ(20.0, tableStyle->get_TopPadding());
    ASSERT_EQ(6, table->get_FirstRow()->get_RowFormat()->get_Borders()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Border>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Border> b)>>([](System::SharedPtr<Aspose::Words::Border> b) -> bool
    {
        return b->get_Color().ToArgb() == System::Drawing::Color::get_Blue().ToArgb();
    }))));
    ASSERT_EQ(Aspose::Words::Tables::CellVerticalAlignment::Center, tableStyle->get_VerticalAlignment());
    
    tableStyle = System::ExplicitCast<Aspose::Words::TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle1"));
    
    ASSERT_TRUE(tableStyle->get_AllowBreakAcrossPages());
    ASSERT_TRUE(tableStyle->get_Bidi());
    ASPOSE_ASSERT_EQ(5.0, tableStyle->get_CellSpacing());
    ASPOSE_ASSERT_EQ(20.0, tableStyle->get_BottomPadding());
    ASPOSE_ASSERT_EQ(5.0, tableStyle->get_LeftPadding());
    ASPOSE_ASSERT_EQ(10.0, tableStyle->get_RightPadding());
    ASPOSE_ASSERT_EQ(20.0, tableStyle->get_TopPadding());
    ASSERT_EQ(System::Drawing::Color::get_AntiqueWhite().ToArgb(), tableStyle->get_Shading()->get_BackgroundPatternColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), tableStyle->get_Borders()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::LineStyle::DotDash, tableStyle->get_Borders()->get_LineStyle());
    ASSERT_EQ(Aspose::Words::Tables::CellVerticalAlignment::Center, tableStyle->get_VerticalAlignment());
}

namespace gtest_test
{

TEST_F(ExTable, TableStyleCreation)
{
    s_instance->TableStyleCreation();
}

} // namespace gtest_test

void ExTable::SetTableAlignment()
{
    //ExStart
    //ExFor:TableStyle.Alignment
    //ExFor:TableStyle.LeftIndent
    //ExSummary:Shows how to set the position of a table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are two ways of aligning a table horizontally.
    // 1 -  Use the "Alignment" property to align it to a location on the page, such as the center:
    auto tableStyle = System::ExplicitCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle1"));
    tableStyle->set_Alignment(Aspose::Words::Tables::TableAlignment::Center);
    tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Blue());
    tableStyle->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::Single);
    
    // Insert a table and apply the style we created to it.
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Aligned to the center of the page");
    builder->EndTable();
    table->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPoints(300));
    
    table->set_Style(tableStyle);
    
    // 2 -  Use the "LeftIndent" to specify an indent from the left margin of the page:
    tableStyle = System::ExplicitCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle2"));
    tableStyle->set_LeftIndent(55);
    tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Green());
    tableStyle->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::Single);
    
    table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Aligned according to left indent");
    builder->EndTable();
    table->set_PreferredWidth(Aspose::Words::Tables::PreferredWidth::FromPoints(300));
    
    table->set_Style(tableStyle);
    
    doc->Save(get_ArtifactsDir() + u"Table.SetTableAlignment.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.SetTableAlignment.docx");
    
    tableStyle = System::ExplicitCast<Aspose::Words::TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle1"));
    
    ASSERT_EQ(Aspose::Words::Tables::TableAlignment::Center, tableStyle->get_Alignment());
    ASPOSE_ASSERT_EQ(tableStyle, doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_Style());
    
    tableStyle = System::ExplicitCast<Aspose::Words::TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle2"));
    
    ASPOSE_ASSERT_EQ(55.0, tableStyle->get_LeftIndent());
    ASPOSE_ASSERT_EQ(tableStyle, (System::ExplicitCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true)))->get_Style());
}

namespace gtest_test
{

TEST_F(ExTable, SetTableAlignment)
{
    s_instance->SetTableAlignment();
}

} // namespace gtest_test

void ExTable::ConditionalStyles()
{
    //ExStart
    //ExFor:ConditionalStyle
    //ExFor:ConditionalStyle.Shading
    //ExFor:ConditionalStyle.Borders
    //ExFor:ConditionalStyle.ParagraphFormat
    //ExFor:ConditionalStyle.BottomPadding
    //ExFor:ConditionalStyle.LeftPadding
    //ExFor:ConditionalStyle.RightPadding
    //ExFor:ConditionalStyle.TopPadding
    //ExFor:ConditionalStyle.Font
    //ExFor:ConditionalStyle.Type
    //ExFor:ConditionalStyleCollection.GetEnumerator
    //ExFor:ConditionalStyleCollection.FirstRow
    //ExFor:ConditionalStyleCollection.LastRow
    //ExFor:ConditionalStyleCollection.LastColumn
    //ExFor:ConditionalStyleCollection.Count
    //ExFor:ConditionalStyleCollection
    //ExFor:ConditionalStyleCollection.BottomLeftCell
    //ExFor:ConditionalStyleCollection.BottomRightCell
    //ExFor:ConditionalStyleCollection.EvenColumnBanding
    //ExFor:ConditionalStyleCollection.EvenRowBanding
    //ExFor:ConditionalStyleCollection.FirstColumn
    //ExFor:ConditionalStyleCollection.Item(ConditionalStyleType)
    //ExFor:ConditionalStyleCollection.Item(Int32)
    //ExFor:ConditionalStyleCollection.OddColumnBanding
    //ExFor:ConditionalStyleCollection.OddRowBanding
    //ExFor:ConditionalStyleCollection.TopLeftCell
    //ExFor:ConditionalStyleCollection.TopRightCell
    //ExFor:ConditionalStyleType
    //ExFor:TableStyle.ConditionalStyles
    //ExSummary:Shows how to work with certain area styles of a table.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Cell 1");
    builder->InsertCell();
    builder->Write(u"Cell 2");
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"Cell 3");
    builder->InsertCell();
    builder->Write(u"Cell 4");
    builder->EndTable();
    
    // Create a custom table style.
    auto tableStyle = System::ExplicitCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle1"));
    
    // Conditional styles are formatting changes that affect only some of the table's cells
    // based on a predicate, such as the cells being in the last row.
    // Below are three ways of accessing a table style's conditional styles from the "ConditionalStyles" collection.
    // 1 -  By style type:
    tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::FirstRow)->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_AliceBlue());
    
    // 2 -  By index:
    tableStyle->get_ConditionalStyles()->idx_get(0)->get_Borders()->set_Color(System::Drawing::Color::get_Black());
    tableStyle->get_ConditionalStyles()->idx_get(0)->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::DotDash);
    ASSERT_EQ(Aspose::Words::ConditionalStyleType::FirstRow, tableStyle->get_ConditionalStyles()->idx_get(0)->get_Type());
    
    // 3 -  As a property:
    tableStyle->get_ConditionalStyles()->get_FirstRow()->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    
    // Apply padding and text formatting to conditional styles.
    tableStyle->get_ConditionalStyles()->get_LastRow()->set_BottomPadding(10);
    tableStyle->get_ConditionalStyles()->get_LastRow()->set_LeftPadding(10);
    tableStyle->get_ConditionalStyles()->get_LastRow()->set_RightPadding(10);
    tableStyle->get_ConditionalStyles()->get_LastRow()->set_TopPadding(10);
    tableStyle->get_ConditionalStyles()->get_LastColumn()->get_Font()->set_Bold(true);
    
    // List all possible style conditions.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::ConditionalStyle>>> enumerator = tableStyle->get_ConditionalStyles()->GetEnumerator();
        while (enumerator->MoveNext())
        {
            System::SharedPtr<Aspose::Words::ConditionalStyle> currentStyle = enumerator->get_Current();
            if (currentStyle != nullptr)
            {
                std::cout << System::EnumGetName(currentStyle->get_Type()) << std::endl;
            }
        }
    }
    
    // Apply the custom style, which contains all conditional styles, to the table.
    table->set_Style(tableStyle);
    
    // Our style applies some conditional styles by default.
    ASSERT_EQ(Aspose::Words::Tables::TableStyleOptions::FirstRow | Aspose::Words::Tables::TableStyleOptions::FirstColumn | Aspose::Words::Tables::TableStyleOptions::RowBands, table->get_StyleOptions());
    
    // We will need to enable all other styles ourselves via the "StyleOptions" property.
    table->set_StyleOptions(table->get_StyleOptions() | Aspose::Words::Tables::TableStyleOptions::LastRow | Aspose::Words::Tables::TableStyleOptions::LastColumn);
    
    doc->Save(get_ArtifactsDir() + u"Table.ConditionalStyles.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.ConditionalStyles.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(Aspose::Words::Tables::TableStyleOptions::Default | Aspose::Words::Tables::TableStyleOptions::LastRow | Aspose::Words::Tables::TableStyleOptions::LastColumn, table->get_StyleOptions());
    System::SharedPtr<Aspose::Words::ConditionalStyleCollection> conditionalStyles = (System::ExplicitCast<Aspose::Words::TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle1")))->get_ConditionalStyles();
    
    ASSERT_EQ(Aspose::Words::ConditionalStyleType::FirstRow, conditionalStyles->idx_get(0)->get_Type());
    ASSERT_EQ(System::Drawing::Color::get_AliceBlue().ToArgb(), conditionalStyles->idx_get(0)->get_Shading()->get_BackgroundPatternColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Black().ToArgb(), conditionalStyles->idx_get(0)->get_Borders()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::LineStyle::DotDash, conditionalStyles->idx_get(0)->get_Borders()->get_LineStyle());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, conditionalStyles->idx_get(0)->get_ParagraphFormat()->get_Alignment());
    
    ASSERT_EQ(Aspose::Words::ConditionalStyleType::LastRow, conditionalStyles->idx_get(2)->get_Type());
    ASPOSE_ASSERT_EQ(10.0, conditionalStyles->idx_get(2)->get_BottomPadding());
    ASPOSE_ASSERT_EQ(10.0, conditionalStyles->idx_get(2)->get_LeftPadding());
    ASPOSE_ASSERT_EQ(10.0, conditionalStyles->idx_get(2)->get_RightPadding());
    ASPOSE_ASSERT_EQ(10.0, conditionalStyles->idx_get(2)->get_TopPadding());
    
    ASSERT_EQ(Aspose::Words::ConditionalStyleType::LastColumn, conditionalStyles->idx_get(3)->get_Type());
    ASSERT_TRUE(conditionalStyles->idx_get(3)->get_Font()->get_Bold());
}

namespace gtest_test
{

TEST_F(ExTable, ConditionalStyles)
{
    s_instance->ConditionalStyles();
}

} // namespace gtest_test

void ExTable::ClearTableStyleFormatting()
{
    //ExStart
    //ExFor:ConditionalStyle.ClearFormatting
    //ExFor:ConditionalStyleCollection.ClearFormatting
    //ExSummary:Shows how to reset conditional table styles.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"First row");
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"Last row");
    builder->EndTable();
    
    auto tableStyle = System::ExplicitCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle1"));
    table->set_Style(tableStyle);
    
    // Set the table style to color the borders of the first row of the table in red.
    tableStyle->get_ConditionalStyles()->get_FirstRow()->get_Borders()->set_Color(System::Drawing::Color::get_Red());
    
    // Set the table style to color the borders of the last row of the table in blue.
    tableStyle->get_ConditionalStyles()->get_LastRow()->get_Borders()->set_Color(System::Drawing::Color::get_Blue());
    
    // Below are two ways of using the "ClearFormatting" method to clear the conditional styles.
    // 1 -  Clear the conditional styles for a specific part of a table:
    tableStyle->get_ConditionalStyles()->idx_get(0)->ClearFormatting();
    
    ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty, tableStyle->get_ConditionalStyles()->get_FirstRow()->get_Borders()->get_Color());
    
    // 2 -  Clear the conditional styles for the entire table:
    tableStyle->get_ConditionalStyles()->ClearFormatting();
    
    ASSERT_TRUE(tableStyle->get_ConditionalStyles()->LINQ_All(static_cast<System::Func<System::SharedPtr<Aspose::Words::ConditionalStyle>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::ConditionalStyle> s)>>([](System::SharedPtr<Aspose::Words::ConditionalStyle> s) -> bool
    {
        return s->get_Borders()->get_Color() == System::Drawing::Color::Empty;
    }))));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, ClearTableStyleFormatting)
{
    s_instance->ClearTableStyleFormatting();
}

} // namespace gtest_test

void ExTable::AlternatingRowStyles()
{
    //ExStart
    //ExFor:TableStyle.ColumnStripe
    //ExFor:TableStyle.RowStripe
    //ExSummary:Shows how to create conditional table styles that alternate between rows.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // We can configure a conditional style of a table to apply a different color to the row/column,
    // based on whether the row/column is even or odd, creating an alternating color pattern.
    // We can also apply a number n to the row/column banding,
    // meaning that the color alternates after every n rows/columns instead of one.
    // Create a table where single columns and rows will band the columns will banded in threes.
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    for (int32_t i = 0; i < 15; i++)
    {
        for (int32_t j = 0; j < 4; j++)
        {
            builder->InsertCell();
            builder->Writeln(System::String::Format(u"{0} column.", (j % 2 == 0 ? System::String(u"Even") : System::String(u"Odd"))));
            builder->Write(System::String::Format(u"Row banding {0}.", (i % 3 == 0 ? System::String(u"start") : System::String(u"continuation"))));
        }
        builder->EndRow();
    }
    builder->EndTable();
    
    // Apply a line style to all the borders of the table.
    auto tableStyle = System::ExplicitCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle1"));
    tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Black());
    tableStyle->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::Double);
    
    // Set the two colors, which will alternate over every 3 rows.
    tableStyle->set_RowStripe(3);
    tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::OddRowBanding)->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightBlue());
    tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::EvenRowBanding)->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightCyan());
    
    // Set a color to apply to every even column, which will override any custom row coloring.
    tableStyle->set_ColumnStripe(1);
    tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::EvenColumnBanding)->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightSalmon());
    
    table->set_Style(tableStyle);
    
    // The "StyleOptions" property enables row banding by default.
    ASSERT_EQ(Aspose::Words::Tables::TableStyleOptions::FirstRow | Aspose::Words::Tables::TableStyleOptions::FirstColumn | Aspose::Words::Tables::TableStyleOptions::RowBands, table->get_StyleOptions());
    
    // Use the "StyleOptions" property also to enable column banding.
    table->set_StyleOptions(table->get_StyleOptions() | Aspose::Words::Tables::TableStyleOptions::ColumnBands);
    
    doc->Save(get_ArtifactsDir() + u"Table.AlternatingRowStyles.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.AlternatingRowStyles.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    tableStyle = System::ExplicitCast<Aspose::Words::TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle1"));
    
    ASPOSE_ASSERT_EQ(tableStyle, table->get_Style());
    ASSERT_EQ(table->get_StyleOptions() | Aspose::Words::Tables::TableStyleOptions::ColumnBands, table->get_StyleOptions());
    
    ASSERT_EQ(System::Drawing::Color::get_Black().ToArgb(), tableStyle->get_Borders()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::LineStyle::Double, tableStyle->get_Borders()->get_LineStyle());
    ASSERT_EQ(3, tableStyle->get_RowStripe());
    ASSERT_EQ(System::Drawing::Color::get_LightBlue().ToArgb(), tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::OddRowBanding)->get_Shading()->get_BackgroundPatternColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_LightCyan().ToArgb(), tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::EvenRowBanding)->get_Shading()->get_BackgroundPatternColor().ToArgb());
    ASSERT_EQ(1, tableStyle->get_ColumnStripe());
    ASSERT_EQ(System::Drawing::Color::get_LightSalmon().ToArgb(), tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::EvenColumnBanding)->get_Shading()->get_BackgroundPatternColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExTable, AlternatingRowStyles)
{
    s_instance->AlternatingRowStyles();
}

} // namespace gtest_test

void ExTable::ConvertToHorizontallyMergedCells()
{
    //ExStart
    //ExFor:Table.ConvertToHorizontallyMergedCells
    //ExSummary:Shows how to convert cells horizontally merged by width to cells merged by CellFormat.HorizontalMerge.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Table with merged cells.docx");
    
    // Microsoft Word does not write merge flags anymore, defining merged cells by width instead.
    // Aspose.Words by default define only 5 cells in a row, and none of them have the horizontal merge flag,
    // even though there were 7 cells in the row before the horizontal merging took place.
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    System::SharedPtr<Aspose::Words::Tables::Row> row = table->get_Rows()->idx_get(0);
    
    ASSERT_EQ(5, row->get_Cells()->get_Count());
    ASSERT_TRUE(row->get_Cells()->LINQ_All(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> c)>>([](System::SharedPtr<Aspose::Words::Node> c) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Tables::Cell>(c))->get_CellFormat()->get_HorizontalMerge() == Aspose::Words::Tables::CellMerge::None;
    }))));
    
    // Use the "ConvertToHorizontallyMergedCells" method to convert cells horizontally merged
    // by its width to the cell horizontally merged by flags.
    // Now, we have 7 cells, and some of them have horizontal merge values.
    table->ConvertToHorizontallyMergedCells();
    row = table->get_Rows()->idx_get(0);
    
    ASSERT_EQ(7, row->get_Cells()->get_Count());
    
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::None, row->get_Cells()->idx_get(0)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::First, row->get_Cells()->idx_get(1)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::Previous, row->get_Cells()->idx_get(2)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::None, row->get_Cells()->idx_get(3)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::First, row->get_Cells()->idx_get(4)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::Previous, row->get_Cells()->idx_get(5)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::None, row->get_Cells()->idx_get(6)->get_CellFormat()->get_HorizontalMerge());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, ConvertToHorizontallyMergedCells)
{
    s_instance->ConvertToHorizontallyMergedCells();
}

} // namespace gtest_test

void ExTable::GetTextFromCells()
{
    //ExStart
    //ExFor:Row.NextRow
    //ExFor:Row.PreviousRow
    //ExFor:Cell.NextCell
    //ExFor:Cell.PreviousCell
    //ExSummary:Shows how to enumerate through all table cells.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    // Enumerate through all cells of the table.
    for (System::SharedPtr<Aspose::Words::Tables::Row> row = table->get_FirstRow(); row != nullptr; row = row->get_NextRow())
    {
        for (System::SharedPtr<Aspose::Words::Tables::Cell> cell = row->get_FirstCell(); cell != nullptr; cell = cell->get_NextCell())
        {
            std::cout << cell->GetText() << std::endl;
        }
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, GetTextFromCells)
{
    s_instance->GetTextFromCells();
}

} // namespace gtest_test

void ExTable::ConvertWithParagraphMark()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Nested tables.docx");
    auto table = System::ExplicitCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    
    // Replace the table with the new paragraph
    ConvertTable(table);
    // Remove table after convertion.
    table->Remove();
    
    doc->Save(get_ArtifactsDir() + u"Table.ConvertWithParagraphMark.docx");
}

namespace gtest_test
{

TEST_F(ExTable, ConvertWithParagraphMark)
{
    s_instance->ConvertWithParagraphMark();
}

} // namespace gtest_test

void ExTable::ConvertWith()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Nested tables.docx");
    auto table = System::ExplicitCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    
    // Convert table to text with specified separator.
    ConvertWith(Aspose::Words::ControlChar::Tab(), table);
    // Remove table after convertion.
    table->Remove();
    
    doc->Save(get_ArtifactsDir() + u"Table.ConvertWith.docx");
}

namespace gtest_test
{

TEST_F(ExTable, ConvertWith)
{
    s_instance->ConvertWith();
}

} // namespace gtest_test

void ExTable::GetColSpanRowSpan()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Merged table.docx");
    
    auto table = System::ExplicitCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    // Convert cells with merged columns into a format that can be easily manipulated.
    table->ConvertToHorizontallyMergedCells();
    
    for (auto&& row : System::IterateOver<Aspose::Words::Tables::Row>(table->get_Rows()))
    {
        auto cell = row->get_FirstCell();
        
        while (cell != nullptr)
        {
            int32_t rowIndex = table->IndexOf(row);
            int32_t cellIndex = cell->get_ParentRow()->IndexOf(cell);
            
            int32_t rowSpan = 1;
            int32_t colSpan = 1;
            
            // Check if the current cell is the start of a vertically merged set of cells.
            if (cell->get_CellFormat()->get_VerticalMerge() == Aspose::Words::Tables::CellMerge::First)
            {
                rowSpan = CalculateRowSpan(table, rowIndex, cellIndex);
            }
            
            // Check if the current cell is the start of a horizontally merged set of cells.
            if (cell->get_CellFormat()->get_HorizontalMerge() == Aspose::Words::Tables::CellMerge::First)
            {
                cell = CalculateColSpan(cell, colSpan);
            }
            else
            {
                cell = cell->get_NextCell();
            }
            
            std::cout << System::String::Format(u"RowIndex = {0}\t ColSpan = {1}\t RowSpan = {2}", rowIndex, colSpan, rowSpan) << std::endl;
        }
    }
}

namespace gtest_test
{

TEST_F(ExTable, GetColSpanRowSpan)
{
    s_instance->GetColSpanRowSpan();
}

} // namespace gtest_test

void ExTable::ContextTableFormatting()
{
    //ExStart:ContextTableFormatting
    //GistId:e06aa7a168b57907a5598e823a22bf0a
    //ExFor:DocumentBuilder.#ctor(Document, DocumentBuilderOptions)
    //ExFor:DocumentBuilder.#ctor(DocumentBuilderOptions)
    //ExFor:DocumentBuilderOptions
    //ExFor:DocumentBuilderOptions.ContextTableFormatting
    //ExSummary:Shows how to ignore table formatting for content after.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builderOptions = System::MakeObject<Aspose::Words::DocumentBuilderOptions>();
    builderOptions->set_ContextTableFormatting(true);
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc, builderOptions);
    
    // Adds content before the table.
    // Default font size is 12.
    builder->Writeln(u"Font size 12 here.");
    builder->StartTable();
    builder->InsertCell();
    // Changes the font size inside the table.
    builder->get_Font()->set_Size(5);
    builder->Write(u"Font size 5 here");
    builder->InsertCell();
    builder->Write(u"Font size 5 here");
    builder->EndRow();
    builder->EndTable();
    
    // If ContextTableFormatting is true, then table formatting isn't applied to the content after.
    // If ContextTableFormatting is false, then table formatting is applied to the content after.
    builder->Writeln(u"Font size 12 here.");
    
    doc->Save(get_ArtifactsDir() + u"Table.ContextTableFormatting.docx");
    //ExEnd:ContextTableFormatting
}

namespace gtest_test
{

TEST_F(ExTable, ContextTableFormatting)
{
    s_instance->ContextTableFormatting();
}

} // namespace gtest_test

void ExTable::AutofitToWindow()
{
    auto expectedPercents = System::MakeArray<double>({51, 49});
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Table wrapped by text.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    table->AutoFit(Aspose::Words::Tables::AutoFitBehavior::AutoFitToWindow);
    
    ASSERT_EQ(expectedPercents->get_Length(), table->get_FirstRow()->get_Cells()->get_Count());
    
    for (auto&& row : System::IterateOver<Aspose::Words::Tables::Row>(table->get_Rows()))
    {
        int32_t i = 0;
        for (auto&& cell : System::IterateOver<Aspose::Words::Tables::Cell>(row->get_Cells()))
        {
            double expectedPercent = expectedPercents[i];
            
            System::SharedPtr<Aspose::Words::Tables::PreferredWidth> cellPrefferedWidth = cell->get_CellFormat()->get_PreferredWidth();
            ASPOSE_ASSERT_EQ(expectedPercent, cellPrefferedWidth->get_Value());
            
            i++;
        }
    }
}

namespace gtest_test
{

TEST_F(ExTable, AutofitToWindow)
{
    s_instance->AutofitToWindow();
}

} // namespace gtest_test

void ExTable::HiddenRow()
{
    //ExStart:HiddenRow
    //GistId:67c1d01ce69d189983b497fd497a7768
    //ExFor:Row.Hidden
    //ExSummary:Shows how to hide a table row.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Row> row = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_FirstRow();
    row->set_Hidden(true);
    
    doc->Save(get_ArtifactsDir() + u"Table.HiddenRow.docx");
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Table.HiddenRow.docx");
    
    row = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_FirstRow();
    ASSERT_TRUE(row->get_Hidden());
    
    for (auto&& cell : System::IterateOver<Aspose::Words::Tables::Cell>(row->get_Cells()))
    {
        for (auto&& para : System::IterateOver<Aspose::Words::Paragraph>(cell->get_Paragraphs()))
        {
            for (auto&& run : System::IterateOver<Aspose::Words::Run>(para->get_Runs()))
            {
                ASSERT_TRUE(run->get_Font()->get_Hidden());
            }
        }
    }
    //ExEnd:HiddenRow
}

namespace gtest_test
{

TEST_F(ExTable, HiddenRow)
{
    s_instance->HiddenRow();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
