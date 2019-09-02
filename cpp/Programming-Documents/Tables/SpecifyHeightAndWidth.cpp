#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include "Model/Document/DocumentBuilder.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Tables/PreferredWidth.h>
#include <Model/Tables/PreferredWidthType.h>
#include <Model/Borders/Shading.h>
#include <Model/Nodes/NodeType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace
{
    void AutoFitToPageWidth(System::String const &outputDataDir)
    {
        // ExStart:AutoFitToPageWidth
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        // Insert a table with a width that takes up half the page width.
        System::SharedPtr<Table> table = builder->StartTable();
        // Insert a few cells
        builder->InsertCell();
        table->set_PreferredWidth(PreferredWidth::FromPercent(50));
        builder->Writeln(u"Cell #1");
        builder->InsertCell();
        builder->Writeln(u"Cell #2");
        builder->InsertCell();
        builder->Writeln(u"Cell #3");
        System::String outputPath = outputDataDir + u"SpecifyHeightAndWidth.AutoFitToPageWidth.doc";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:AutoFitToPageWidth
        std::cout << "Table autofit successfully to 50% of the page width." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetPreferredWidthSettings(System::String const &outputDataDir)
    {
        // ExStart:SetPreferredWidthSettings
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        // Insert a table row made up of three cells which have different preferred widths.
        System::SharedPtr<Table> table = builder->StartTable();
        // Insert an absolute sized cell.
        builder->InsertCell();
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPoints(40));
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightYellow());
        builder->Writeln(u"Cell at 40 points width");
        // Insert a relative (percent) sized cell.
        builder->InsertCell();
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(20));
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightBlue());
        builder->Writeln(u"Cell at 20% width");
        // Insert a auto sized cell.
        builder->InsertCell();
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::Auto());
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightGreen());
        builder->Writeln(u"Cell automatically sized. The size of this cell is calculated from the table preferred width.");
        builder->Writeln(u"In this case the cell will fill up the rest of the available space.");
        System::String outputPath = outputDataDir + u"SpecifyHeightAndWidth.SetPreferredWidthSettings.doc";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetPreferredWidthSettings
        std::cout << "Different preferred width settings set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void RetrievePreferredWidthType(System::String const &inputDataDir)
    {
        // ExStart:RetrievePreferredWidthType
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.SimpleTable.doc");
        // Retrieve the first table in the document.
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // ExStart:AllowAutoFit
        table->set_AllowAutoFit(true);
        // ExEnd:AllowAutoFit
        System::SharedPtr<Cell> firstCell = table->get_FirstRow()->get_FirstCell();
        PreferredWidthType type = firstCell->get_CellFormat()->get_PreferredWidth()->get_Type();
        double value = firstCell->get_CellFormat()->get_PreferredWidth()->get_Value();
        // ExEnd:RetrievePreferredWidthType
        std::cout << "Table preferred width type value is " << value << std::endl;
    }
}

void SpecifyHeightAndWidth()
{
    std::cout << "SpecifyHeightAndWidth example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithTables();
    System::String outputDataDir = GetOutputDataDir_WorkingWithTables();
    AutoFitToPageWidth(outputDataDir);
    SetPreferredWidthSettings(outputDataDir);
    RetrievePreferredWidthType(inputDataDir);
    std::cout << "SpecifyHeightAndWidth example finished." << std::endl << std::endl;
}