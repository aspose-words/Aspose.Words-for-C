#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/test_tools/method_argument_tuple.h>
#include <cstdint>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Tables { class Table; } } }
namespace Aspose { namespace Words { namespace Drawing { class Shape; } } }
namespace Aspose { namespace Words { class Document; } }
namespace Aspose { namespace Words { namespace Tables { class Cell; } } }

namespace ApiExamples {

/// <summary>
/// Examples using tables in documents.
/// </summary>
class ExTable : public ApiExampleBase
{
public:

    void CreateTable();
    void RowCellFormat();
    void DisplayContentOfTables();
    void CalculateDepthOfNestedTables();
    void ConvertTextBoxToTable();
    void EnsureTableMinimum();
    void EnsureRowMinimum();
    void EnsureCellMinimum();
    void SetOutlineBorders();
    void SetTableBorders();
    void RowFormat();
    void CellFormat();
    void GetDistance();
    void Borders();
    void ReplaceCellText();
    void PrintTableRange();
    void CloneTable();
    void DisableBreakAcrossPages();
    void AllowAutoFitOnTable(bool allowAutoFit);
    void KeepTableTogether();
    void FixDefaultTableWidthsInAw105();
    void FixDefaultTableBordersIn105();
    void FixDefaultTableFormattingExceptionIn105();
    void FixRowFormattingNotAppliedIn105();
    void GetIndexOfTableElements();
    void GetPreferredWidthTypeAndValue();
    void AllowCellSpacing(bool allowCellSpacing);
    void CreateNestedTable();
    void CheckCellsMerged();
    System::String PrintCellMergeType(System::SharedPtr<Aspose::Words::Tables::Cell> cell);
    void MergeCellRange();
    /// <summary>
    /// Merges the range of cells found between the two specified cells both horizontally and vertically. Can span over multiple rows.
    /// </summary>
    static void MergeCells(System::SharedPtr<Aspose::Words::Tables::Cell> startCell, System::SharedPtr<Aspose::Words::Tables::Cell> endCell);
    void CombineTables();
    void SplitTable();
    void WrapText();
    void GetFloatingTableProperties();
    void ChangeFloatingTableProperties();
    void TableStyleCreation();
    void SetTableAlignment();
    void ConditionalStyles();
    void ClearTableStyleFormatting();
    void AlternatingRowStyles();
    void ConvertToHorizontallyMergedCells();
    
protected:

    /// <summary>
    /// Calculates what level a table is nested inside other tables.
    /// <returns>
    /// An integer containing the level the table is nested at.
    /// 0 = Table is not nested inside any other table
    /// 1 = Table is nested within one parent table
    /// 2 = Table is nested within two parent tables etc..</returns>
    /// </summary>
    static int32_t GetNestedDepthOfTable(System::SharedPtr<Aspose::Words::Tables::Table> table);
    /// <summary>
    /// Determines if a table contains any immediate child table within its cells.
    /// Does not recursively traverse through those tables to check for further tables.
    /// <returns>Returns true if at least one child cell contains a table.
    /// Returns false if no cells in the table contains a table.</returns>
    /// </summary>
    static int32_t GetChildTableCount(System::SharedPtr<Aspose::Words::Tables::Table> table);
    /// <summary>
    /// Converts a textbox to a table by copying the same content and formatting.
    /// Currently export to HTML will render the textbox as an image which looses any text functionality.
    /// This is useful to convert textboxes in order to retain proper text.
    /// </summary>
    /// <param name="textBox">The textbox shape to convert to a table</param>
    static void ConvertTextboxToTable(System::SharedPtr<Aspose::Words::Drawing::Shape> textBox);
    /// <summary>
    /// Creates a new table in the document with the given dimensions and text in each cell.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Tables::Table> CreateTable(System::SharedPtr<Aspose::Words::Document> doc, int32_t rowCount, int32_t cellCount, System::String cellText);
    void TestCreateNestedTable(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples


