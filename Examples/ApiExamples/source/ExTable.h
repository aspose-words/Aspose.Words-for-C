#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Tables;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExTable : public ApiExampleBase
{
    typedef ExTable ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void CreateTable();
    void Padding();
    void RowCellFormat();
    void DisplayContentOfTables();
    void EnsureTableMinimum();
    void EnsureRowMinimum();
    void EnsureCellMinimum();
    void SetOutlineBorders();
    void SetBorders();
    void RowFormat();
    void CellFormat();
    void DistanceBetweenTableAndText();
    void Borders();
    void ReplaceCellText();
    void RemoveParagraphTextAndMark(bool isSmartParagraphBreakReplacement);
    void PrintTableRange();
    void CloneTable();
    void AllowBreakAcrossPages(bool allowBreakAcrossPages);
    void AllowAutoFitOnTable(bool allowAutoFit);
    void KeepTableTogether();
    void GetIndexOfTableElements();
    void GetPreferredWidthTypeAndValue();
    void AllowCellSpacing(bool allowCellSpacing);
    void CreateNestedTable();
    void CheckCellsMerged();
    System::String PrintCellMergeType(System::SharedPtr<Aspose::Words::Tables::Cell> cell);
    void MergeCellRange();
    /// <summary>
    /// Merges the range of cells found between the two specified cells both horizontally and vertically.
    /// Can span over multiple rows.
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
    void GetTextFromCells();
    void ConvertWithParagraphMark();
    void ConvertWith();
    void GetColSpanRowSpan();
    void ContextTableFormatting();
    void AutofitToWindow();
    void HiddenRow();
    
protected:

    /// <summary>
    /// Determines if a table contains any immediate child table within its cells.
    /// Do not recursively traverse through those tables to check for further tables.
    /// </summary>
    /// <returns>
    /// Returns true if at least one child cell contains a table.
    /// Returns false if no cells in the table contain a table.
    /// </returns>
    static int32_t GetChildTableCount(System::SharedPtr<Aspose::Words::Tables::Table> table);
    /// <summary>
    /// Creates a new table in the document with the given dimensions and text in each cell.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Tables::Table> CreateTable(System::SharedPtr<Aspose::Words::Document> doc, int32_t rowCount, int32_t cellCount, System::String cellText);
    void TestCreateNestedTable(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Recursively converts nested tables within a given table.
    /// </summary>
    /// <param name="table">The table to be converted.</param>
    void ConvertTable(System::SharedPtr<Aspose::Words::Tables::Table> table);
    /// <summary>
    /// Converts the content of a table into a series of paragraphs, separated by a specified separator.
    /// </summary>
    /// <param name="separator">The string used to separate the content of each cell.</param>
    /// <param name="table">The table to be converted.</param>
    void ConvertWith(System::String separator, System::SharedPtr<Aspose::Words::Tables::Table> table);
    /// <summary>
    /// Calculates the row span for a cell in a table.
    /// </summary>
    /// <param name="table">The table containing the cell.</param>
    /// <param name="rowIndex">The index of the row containing the cell.</param>
    /// <param name="cellIndex">The index of the cell within the row.</param>
    /// <returns>The number of rows spanned by the cell.</returns>
    int32_t CalculateRowSpan(System::SharedPtr<Aspose::Words::Tables::Table> table, int32_t rowIndex, int32_t cellIndex);
    /// <summary>
    /// Calculates the column span of a cell based on its horizontal merge settings.
    /// </summary>
    /// <param name="cell">The cell for which to calculate the column span.</param>
    /// <param name="colSpan">The resulting column span value.</param>
    /// <returns>The next cell in the sequence after calculating the column span.</returns>
    System::SharedPtr<Aspose::Words::Tables::Cell> CalculateColSpan(System::SharedPtr<Aspose::Words::Tables::Cell> cell, int32_t& colSpan);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


