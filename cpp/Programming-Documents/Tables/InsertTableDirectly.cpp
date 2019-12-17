#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Tables/Table.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/RowFormat.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Tables/AutoFitBehavior.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/Body.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/Run.h>
#include <Model/Borders/Shading.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

void InsertTableDirectly()
{
    std::cout << "InsertTableDirectly example started." << std::endl;
    // ExStart:InsertTableDirectly
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithTables();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    // We start by creating the table object. Note how we must pass the document object
    // To the constructor of each node. This is because every node we create must belong
    // To some document.
    System::SharedPtr<Table> table = System::MakeObject<Table>(doc);
    // Add the table to the document.
    doc->get_FirstSection()->get_Body()->AppendChild(table);
    // Here we could call EnsureMinimum to create the rows and cells for us. This method is used
    // To ensure that the specified node is valid, in this case a valid table should have at least one
    // Row and one cell, therefore this method creates them for us.
    // Instead we will handle creating the row and table ourselves. This would be the best way to do this
    // If we were creating a table inside an algorthim for example.
    System::SharedPtr<Row> row = System::MakeObject<Row>(doc);
    row->get_RowFormat()->set_AllowBreakAcrossPages(true);
    table->AppendChild(row);
    // We can now apply any auto fit settings.
    table->AutoFit(AutoFitBehavior::FixedColumnWidths);
    // Create a cell and add it to the row
    System::SharedPtr<Cell> cell = System::MakeObject<Cell>(doc);
    cell->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightBlue());
    cell->get_CellFormat()->set_Width(80);
    // Add a paragraph to the cell as well as a new run with some text.
    cell->AppendChild(System::MakeObject<Paragraph>(doc));
    cell->get_FirstParagraph()->AppendChild(System::MakeObject<Run>(doc, u"Row 1, Cell 1 Text"));
    // Add the cell to the row.
    row->AppendChild(cell);
    // We would then repeat the process for the other cells and rows in the table.
    // We can also speed things up by cloning existing cells and rows.
    row->AppendChild((System::StaticCast<Node>(cell))->Clone(false));
    row->get_LastCell()->AppendChild(System::MakeObject<Paragraph>(doc));
    row->get_LastCell()->get_FirstParagraph()->AppendChild(System::MakeObject<Run>(doc, u"Row 1, Cell 2 Text"));
    System::String outputPath = outputDataDir + u"InsertTableDirectly.doc";
    // Save the document to disk.
    doc->Save(outputPath);
    // ExEnd:InsertTableDirectly
    std::cout << "Table using notes inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertTableDirectly example finished." << std::endl << std::endl;

}