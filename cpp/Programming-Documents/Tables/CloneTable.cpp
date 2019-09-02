#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include "Model/Document/Document.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/CellCollection.h>
#include <Model/Tables/Cell.h>
#include <Model/Text/Paragraph.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/Node.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace
{
    void CloneCompleteTable(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:CloneCompleteTable
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.SimpleTable.doc");
        // Retrieve the first table in the document.
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // Create a clone of the table.
        System::SharedPtr<Table> tableClone = System::DynamicCast<Table>((System::StaticCast<Node>(table))->Clone(true));
        // Insert the cloned table into the document after the original
        table->get_ParentNode()->InsertAfter(tableClone, table);
        // Insert an empty paragraph between the two tables or else they will be combined into one
        // Upon save. This has to do with document validation.
        table->get_ParentNode()->InsertAfter(System::MakeObject<Paragraph>(doc), table);
        System::String outputPath = outputDataDir + u"CloneTable.CloneCompleteTable.doc";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:CloneCompleteTable
        std::cout << "Table cloned successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void CloneLastRow(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:CloneLastRow
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.SimpleTable.doc");
        // Retrieve the first table in the document.
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // Clone the last row in the table.
        System::SharedPtr<Row> clonedRow = System::DynamicCast<Row>((System::StaticCast<Node>(table->get_LastRow()))->Clone(true));
        // Remove all content from the cloned row's cells. This makes the row ready for
        // New content to be inserted into.
        for (System::SharedPtr<Cell> cell : System::IterateOver<System::SharedPtr<Cell>>(clonedRow->get_Cells()))
        {
            cell->RemoveAllChildren();
        }
        // Add the row to the end of the table.
        table->AppendChild(clonedRow);
        System::String outputPath = outputDataDir + u"CloneTable.CloneLastRow.doc";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:CloneLastRow
        std::cout << "Table last row cloned successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void CloneTable()
{
    std::cout << "CloneTable example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithTables();
    System::String outputDataDir = GetOutputDataDir_WorkingWithTables();
    CloneCompleteTable(inputDataDir, outputDataDir);
    CloneLastRow(inputDataDir, outputDataDir);
    std::cout << "CloneTable example finished." << std::endl << std::endl;
}