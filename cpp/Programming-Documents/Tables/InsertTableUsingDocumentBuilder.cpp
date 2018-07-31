#include "stdafx.h"
#include "examples.h"

#include <system/shared_ptr.h>
#include <drawing/color.h>

#include "Model/Document/Document.h"
#include "Model/Document/DocumentBuilder.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/RowFormat.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Tables/CellVerticalAlignment.h>
#include <Model/Text/HeightRule.h>
#include <Model/Text/ParagraphAlignment.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/Font.h>
#include <Model/Borders/Shading.h>

using namespace Aspose::Words;

namespace
{
    void SimpleTable(System::String const &dataDir)
    {
        // ExStart:SimpleTable
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        // We call this method to start building the table.
        builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 1 Content.");
        // Build the second cell
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 2 Content.");
        // Call the following method to end the row and start a new row.
        builder->EndRow();
        // Build the first cell of the second row.
        builder->InsertCell();
        builder->Write(u"Row 2, Cell 1 Content");
        // Build the second cell.
        builder->InsertCell();
        builder->Write(u"Row 2, Cell 2 Content.");
        builder->EndRow();
        // Signal that we have finished building the table.
        builder->EndTable();
        System::String outputPath = dataDir + GetOutputFilePath(u"InsertTableUsingDocumentBuilder.SimpleTable.doc");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SimpleTable
        std::cout << "\nSimple table created successfully.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void FormattedTable(System::String const &dataDir)
    {
        // ExStart:FormattedTable
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        auto table = builder->StartTable();
        // Make the header row.
        builder->InsertCell();
        // Set the left indent for the table. Table wide formatting must be applied after
        // At least one row is present in the table.
        table->set_LeftIndent(20.0);
        // Set height and define the height rule for the header row.
        builder->get_RowFormat()->set_Height(40.0);
        builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::AtLeast);
        // Some special features for the header row.
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::FromArgb(198, 217, 241));
        builder->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
        builder->get_Font()->set_Size(16);
        builder->get_Font()->set_Name(u"Arial");
        builder->get_Font()->set_Bold(true);
        builder->get_CellFormat()->set_Width(100.0);
        builder->Write(u"Header Row,\n Cell 1");
        // We don't need to specify the width of this cell because it's inherited from the previous cell.
        builder->InsertCell();
        builder->Write(u"Header Row,\n Cell 2");
        builder->InsertCell();
        builder->get_CellFormat()->set_Width(200.0);
        builder->Write(u"Header Row,\n Cell 3");
        builder->EndRow();
        // Set features for the other rows and cells.
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_White());
        builder->get_CellFormat()->set_Width(100.0);
        builder->get_CellFormat()->set_VerticalAlignment(Aspose::Words::Tables::CellVerticalAlignment::Center);
        // Reset height and define a different height rule for table body
        builder->get_RowFormat()->set_Height(30.0);
        builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Auto);
        builder->InsertCell();
        // Reset font formatting.
        builder->get_Font()->set_Size(12);
        builder->get_Font()->set_Bold(false);
        // Build the other cells.
        builder->Write(u"Row 1, Cell 1 Content");
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 2 Content");
        builder->InsertCell();
        builder->get_CellFormat()->set_Width(200.0);
        builder->Write(u"Row 1, Cell 3 Content");
        builder->EndRow();
        builder->InsertCell();
        builder->get_CellFormat()->set_Width(100.0);
        builder->Write(u"Row 2, Cell 1 Content");
        builder->InsertCell();
        builder->Write(u"Row 2, Cell 2 Content");
        builder->InsertCell();
        builder->get_CellFormat()->set_Width(200.0);
        builder->Write(u"Row 2, Cell 3 Content.");
        builder->EndRow();
        builder->EndTable();
        System::String outputPath = dataDir + GetOutputFilePath(u"InsertTableUsingDocumentBuilder.FormattedTable.doc");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:FormattedTable
        std::cout << "\nFormatted table created successfully.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void NestedTable(System::String dataDir)
    {
        // ExStart:NestedTable
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        // Build the outer table.
        auto cell = builder->InsertCell();
        builder->Writeln(u"Outer Table Cell 1");
        builder->InsertCell();
        builder->Writeln(u"Outer Table Cell 2");
        // This call is important in order to create a nested table within the first table
        // Without this call the cells inserted below will be appended to the outer table.
        builder->EndTable();
        // Move to the first cell of the outer table.
        builder->MoveTo(cell->get_FirstParagraph());
        // Build the inner table.
        builder->InsertCell();
        builder->Writeln(u"Inner Table Cell 1");
        builder->InsertCell();
        builder->Writeln(u"Inner Table Cell 2");
        builder->EndTable();
        System::String outputPath = dataDir + GetOutputFilePath(u"InsertTableUsingDocumentBuilder.NestedTable.doc");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:NestedTable
        std::cout << "\nNested table created successfully.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void InsertTableUsingDocumentBuilder()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithTables();
    SimpleTable(dataDir);
    FormattedTable(dataDir);
    NestedTable(dataDir);
}