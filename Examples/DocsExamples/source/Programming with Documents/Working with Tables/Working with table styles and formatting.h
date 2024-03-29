#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Border.h>
#include <Aspose.Words.Cpp/BorderCollection.h>
#include <Aspose.Words.Cpp/BorderType.h>
#include <Aspose.Words.Cpp/ConditionalStyle.h>
#include <Aspose.Words.Cpp/ConditionalStyleCollection.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/HeightRule.h>
#include <Aspose.Words.Cpp/LineStyle.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Shading.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/StyleType.h>
#include <Aspose.Words.Cpp/TableStyle.h>
#include <Aspose.Words.Cpp/Tables/AutoFitBehavior.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/Tables/TableAlignment.h>
#include <Aspose.Words.Cpp/Tables/TableStyleOptions.h>
#include <Aspose.Words.Cpp/TextOrientation.h>
#include <Aspose.Words.Cpp/TextureIndex.h>
#include <drawing/color.h>
#include <system/enum_helpers.h>
#include <system/exceptions.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Tables {

class WorkingWithTableStylesAndFormatting : public DocsExamplesBase
{
public:
    void DistanceBetweenTableSurroundingText()
    {
        //ExStart:DistanceBetweenTableSurroundingText
        //GistId:eb66dfc4c4820add33be9df57ba4c4cd
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        std::cout << "\nGet distance between table left, right, bottom, top and the surrounding text." << std::endl;
        auto table = System::ExplicitCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        std::cout << table->get_DistanceTop() << std::endl;
        std::cout << table->get_DistanceBottom() << std::endl;
        std::cout << table->get_DistanceRight() << std::endl;
        std::cout << table->get_DistanceLeft() << std::endl;
        //ExEnd:DistanceBetweenTableSurroundingText
    }

    void ApplyOutlineBorder()
    {
        //ExStart:ApplyOutlineBorder
        //GistId:4f1d7039f19f9f49472a50cc0d0fc475
        //ExStart:InlineTablePosition
        //GistId:eb66dfc4c4820add33be9df57ba4c4cd
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::ExplicitCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // Align the table to the center of the page.
        table->set_Alignment(TableAlignment::Center);
        //ExEnd:InlineTablePosition
        // Clear any existing borders from the table.
        table->ClearBorders();

        // Set a green border around the table but not inside.
        table->SetBorder(BorderType::Left, LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
        table->SetBorder(BorderType::Right, LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
        table->SetBorder(BorderType::Top, LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
        table->SetBorder(BorderType::Bottom, LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);

        // Fill the cells with a light green solid color.
        table->SetShading(TextureIndex::TextureSolid, System::Drawing::Color::get_LightGreen(), System::Drawing::Color::Empty);

        doc->Save(ArtifactsDir + u"WorkingWithTableStylesAndFormatting.ApplyOutlineBorder.docx");
        //ExEnd:ApplyOutlineBorder
    }

    void BuildTableWithBorders()
    {
        //ExStart:BuildTableWithBorders
        //GistId:4f1d7039f19f9f49472a50cc0d0fc475
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::ExplicitCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        // Clear any existing borders from the table.
        table->ClearBorders();

        // Set a green border around and inside the table.
        table->SetBorders(LineStyle::Single, 1.5, System::Drawing::Color::get_Green());

        doc->Save(ArtifactsDir + u"WorkingWithTableStylesAndFormatting.BuildTableWithBorders.docx");
        //ExEnd:BuildTableWithBorders
    }

    void ModifyRowFormatting()
    {
        //ExStart:ModifyRowFormatting
        //GistId:4f1d7039f19f9f49472a50cc0d0fc475
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::ExplicitCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        // Retrieve the first row in the table.
        SharedPtr<Row> firstRow = table->get_FirstRow();
        firstRow->get_RowFormat()->get_Borders()->set_LineStyle(LineStyle::None);
        firstRow->get_RowFormat()->set_HeightRule(HeightRule::Auto);
        firstRow->get_RowFormat()->set_AllowBreakAcrossPages(true);
        //ExEnd:ModifyRowFormatting
    }

    void ApplyRowFormatting()
    {
        //ExStart:ApplyRowFormatting
        //GistId:4f1d7039f19f9f49472a50cc0d0fc475
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();

        SharedPtr<RowFormat> rowFormat = builder->get_RowFormat();
        rowFormat->set_Height(100);
        rowFormat->set_HeightRule(HeightRule::Exactly);

        // These formatting properties are set on the table and are applied to all rows in the table.
        table->set_LeftPadding(30);
        table->set_RightPadding(30);
        table->set_TopPadding(30);
        table->set_BottomPadding(30);

        builder->Writeln(u"I'm a wonderful formatted row.");

        builder->EndRow();
        builder->EndTable();

        doc->Save(ArtifactsDir + u"WorkingWithTableStylesAndFormatting.ApplyRowFormatting.docx");
        //ExEnd:ApplyRowFormatting
    }

    void CellPadding()
    {
        //ExStart:CellPadding
        //GistId:4f1d7039f19f9f49472a50cc0d0fc475
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartTable();
        builder->InsertCell();

        // Sets the amount of space (in points) to add to the left/top/right/bottom of the cell's contents.
        builder->get_CellFormat()->SetPaddings(30, 50, 30, 50);
        builder->Writeln(u"I'm a wonderful formatted cell.");

        builder->EndRow();
        builder->EndTable();

        doc->Save(ArtifactsDir + u"WorkingWithTableStylesAndFormatting.CellPadding.docx");
        //ExEnd:CellPadding
    }

    /// <summary>
    /// Shows how to modify formatting of a table cell.
    /// </summary>
    void ModifyCellFormatting()
    {
        //ExStart:ModifyCellFormatting
        //GistId:4f1d7039f19f9f49472a50cc0d0fc475
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
        auto table = System::ExplicitCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        SharedPtr<Cell> firstCell = table->get_FirstRow()->get_FirstCell();
        firstCell->get_CellFormat()->set_Width(30);
        firstCell->get_CellFormat()->set_Orientation(TextOrientation::Downward);
        firstCell->get_CellFormat()->get_Shading()->set_ForegroundPatternColor(System::Drawing::Color::get_LightGreen());
        //ExEnd:ModifyCellFormatting
    }

    void FormatTableAndCellWithDifferentBorders()
    {
        //ExStart:FormatTableAndCellWithDifferentBorders
        //GistId:4f1d7039f19f9f49472a50cc0d0fc475
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();

        // Set the borders for the entire table.
        table->SetBorders(LineStyle::Single, 2.0, System::Drawing::Color::get_Black());

        // Set the cell shading for this cell.
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Red());
        builder->Writeln(u"Cell #1");

        builder->InsertCell();

        // Specify a different cell shading for the second cell.
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Green());
        builder->Writeln(u"Cell #2");

        builder->EndRow();

        // Clear the cell formatting from previous operations.
        builder->get_CellFormat()->ClearFormatting();

        builder->InsertCell();

        // Create larger borders for the first cell of this row. This will be different
        // compared to the borders set for the table.
        builder->get_CellFormat()->get_Borders()->get_Left()->set_LineWidth(4.0);
        builder->get_CellFormat()->get_Borders()->get_Right()->set_LineWidth(4.0);
        builder->get_CellFormat()->get_Borders()->get_Top()->set_LineWidth(4.0);
        builder->get_CellFormat()->get_Borders()->get_Bottom()->set_LineWidth(4.0);
        builder->Writeln(u"Cell #3");

        builder->InsertCell();
        builder->get_CellFormat()->ClearFormatting();
        builder->Writeln(u"Cell #4");

        doc->Save(ArtifactsDir + u"WorkingWithTableStylesAndFormatting.FormatTableAndCellWithDifferentBorders.docx");
        //ExEnd:FormatTableAndCellWithDifferentBorders
    }

    void TableTitleAndDescription()
    {
        //ExStart:TableTitleAndDescription
        //GistId:0c6e4ebd55874aa8634f1d74d5ef492f
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::ExplicitCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        table->set_Title(u"Test title");
        table->set_Description(u"Test description");

        auto options = MakeObject<OoxmlSaveOptions>();
        options->set_Compliance(OoxmlCompliance::Iso29500_2008_Strict);

        doc->get_CompatibilityOptions()->OptimizeFor(Settings::MsWordVersion::Word2016);

        doc->Save(ArtifactsDir + u"WorkingWithTableStylesAndFormatting.TableTitleAndDescription.docx", options);
        //ExEnd:TableTitleAndDescription
    }

    void AllowCellSpacing()
    {
        //ExStart:AllowCellSpacing
        //GistId:4f1d7039f19f9f49472a50cc0d0fc475
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::ExplicitCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        table->set_AllowCellSpacing(true);
        table->set_CellSpacing(2);

        doc->Save(ArtifactsDir + u"WorkingWithTableStylesAndFormatting.AllowCellSpacing.docx");
        //ExEnd:AllowCellSpacing
    }

    void BuildTableWithStyle()
    {
        //ExStart:BuildTableWithStyle
        //GistId:aeee44b29c42f9e9404ce57a80305383
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();

        // We must insert at least one row first before setting any table formatting.
        builder->InsertCell();

        // Set the table style used based on the unique style identifier.
        table->set_StyleIdentifier(StyleIdentifier::MediumShading1Accent1);

        // Apply which features should be formatted by the style.
        table->set_StyleOptions(TableStyleOptions::FirstColumn | TableStyleOptions::RowBands | TableStyleOptions::FirstRow);
        table->AutoFit(AutoFitBehavior::AutoFitToContents);

        builder->Writeln(u"Item");
        builder->get_CellFormat()->set_RightPadding(40);
        builder->InsertCell();
        builder->Writeln(u"Quantity (kg)");
        builder->EndRow();

        builder->InsertCell();
        builder->Writeln(u"Apples");
        builder->InsertCell();
        builder->Writeln(u"20");
        builder->EndRow();

        builder->InsertCell();
        builder->Writeln(u"Bananas");
        builder->InsertCell();
        builder->Writeln(u"40");
        builder->EndRow();

        builder->InsertCell();
        builder->Writeln(u"Carrots");
        builder->InsertCell();
        builder->Writeln(u"50");
        builder->EndRow();

        doc->Save(ArtifactsDir + u"WorkingWithTableStylesAndFormatting.BuildTableWithStyle.docx");
        //ExEnd:BuildTableWithStyle
    }

    void ExpandFormattingOnCellsAndRowFromStyle()
    {
        //ExStart:ExpandFormattingOnCellsAndRowFromStyle
        //GistId:aeee44b29c42f9e9404ce57a80305383
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        // Get the first cell of the first table in the document.
        auto table = System::ExplicitCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        SharedPtr<Cell> firstCell = table->get_FirstRow()->get_FirstCell();

        // First print the color of the cell shading.
        // This should be empty as the current shading is stored in the table style.
        System::Drawing::Color cellShadingBefore = firstCell->get_CellFormat()->get_Shading()->get_BackgroundPatternColor();
        std::cout << (String(u"Cell shading before style expansion: ") + cellShadingBefore) << std::endl;

        doc->ExpandTableStylesToDirectFormatting();

        // Now print the cell shading after expanding table styles.
        // A blue background pattern color should have been applied from the table style.
        System::Drawing::Color cellShadingAfter = firstCell->get_CellFormat()->get_Shading()->get_BackgroundPatternColor();
        std::cout << (String(u"Cell shading after style expansion: ") + cellShadingAfter) << std::endl;
        //ExEnd:ExpandFormattingOnCellsAndRowFromStyle
    }

    void CreateTableStyle()
    {
        //ExStart:CreateTableStyle
        //GistId:aeee44b29c42f9e9404ce57a80305383
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Name");
        builder->InsertCell();
        builder->Write(u"Value");
        builder->EndRow();
        builder->InsertCell();
        builder->InsertCell();
        builder->EndTable();

        auto tableStyle = System::ExplicitCast<TableStyle>(doc->get_Styles()->Add(StyleType::Table, u"MyTableStyle1"));
        tableStyle->get_Borders()->set_LineStyle(LineStyle::Double);
        tableStyle->get_Borders()->set_LineWidth(1);
        tableStyle->set_LeftPadding(18);
        tableStyle->set_RightPadding(18);
        tableStyle->set_TopPadding(12);
        tableStyle->set_BottomPadding(12);

        table->set_Style(tableStyle);

        doc->Save(ArtifactsDir + u"WorkingWithTableStylesAndFormatting.CreateTableStyle.docx");
        //ExEnd:CreateTableStyle
    }

    void DefineConditionalFormatting()
    {
        //ExStart:DefineConditionalFormatting
        //GistId:aeee44b29c42f9e9404ce57a80305383
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Name");
        builder->InsertCell();
        builder->Write(u"Value");
        builder->EndRow();
        builder->InsertCell();
        builder->InsertCell();
        builder->EndTable();

        auto tableStyle = System::ExplicitCast<TableStyle>(doc->get_Styles()->Add(StyleType::Table, u"MyTableStyle1"));
        tableStyle->get_ConditionalStyles()->get_FirstRow()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_GreenYellow());
        tableStyle->get_ConditionalStyles()->get_FirstRow()->get_Shading()->set_Texture(TextureIndex::TextureNone);

        table->set_Style(tableStyle);

        doc->Save(ArtifactsDir + u"WorkingWithTableStylesAndFormatting.DefineConditionalFormatting.docx");
        //ExEnd:DefineConditionalFormatting
    }

    void SetTableCellFormatting()
    {
        //ExStart:DocumentBuilderSetTableCellFormatting
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartTable();
        builder->InsertCell();

        SharedPtr<CellFormat> cellFormat = builder->get_CellFormat();
        cellFormat->set_Width(250);
        cellFormat->set_LeftPadding(30);
        cellFormat->set_RightPadding(30);
        cellFormat->set_TopPadding(30);
        cellFormat->set_BottomPadding(30);

        builder->Writeln(u"I'm a wonderful formatted cell.");

        builder->EndRow();
        builder->EndTable();

        doc->Save(ArtifactsDir + u"WorkingWithTableStylesAndFormatting.DocumentBuilderSetTableCellFormatting.docx");
        //ExEnd:DocumentBuilderSetTableCellFormatting
    }

    void SetTableRowFormatting()
    {
        //ExStart:DocumentBuilderSetTableRowFormatting
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();

        SharedPtr<RowFormat> rowFormat = builder->get_RowFormat();
        rowFormat->set_Height(100);
        rowFormat->set_HeightRule(HeightRule::Exactly);

        // These formatting properties are set on the table and are applied to all rows in the table.
        table->set_LeftPadding(30);
        table->set_RightPadding(30);
        table->set_TopPadding(30);
        table->set_BottomPadding(30);

        builder->Writeln(u"I'm a wonderful formatted row.");

        builder->EndRow();
        builder->EndTable();

        doc->Save(ArtifactsDir + u"WorkingWithTableStylesAndFormatting.DocumentBuilderSetTableRowFormatting.docx");
        //ExEnd:DocumentBuilderSetTableRowFormatting
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Tables
