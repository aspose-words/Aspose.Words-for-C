#include "stdafx.h"
#include "examples.h"

#include <system/shared_ptr.h>

#include "Model/Document/Document.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/RowCollection.h>
#include <Model/Text/Paragraph.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/NodeType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace
{
    void CombineRows(System::String const &dataDir)
    {
        // ExStart:CombineRows
        // Load the document.
        auto doc = System::MakeObject<Document>(dataDir + u"Table.Document.doc");
        // Get the first and second table in the document.
        // The rows from the second table will be appended to the end of the first table.
        auto firstTable = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        auto secondTable = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 1, true));
        // Append all rows from the current table to the next.
        // Due to the design of tables even tables with different cell count and widths can be joined into one table.
        while (secondTable->get_HasChildNodes())
        {
            firstTable->get_Rows()->Add(secondTable->get_FirstRow());
        }
        // Remove the empty table container.
        secondTable->Remove();
        System::String outputPath = dataDir + GetOutputFilePath(u"JoiningAndSplittingTable.CombineRows.doc");
        // Save the finished document.
        doc->Save(outputPath);
        // ExEnd:CombineRows
        std::cout << "Rows combine successfully from two tables into one." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SplitTable(System::String const &dataDir)
    {
        // ExStart:SplitTable
        // Load the document.
        auto doc = System::MakeObject<Document>(dataDir + u"Table.Document.doc");
        // Get the first table in the document.
        auto firstTable = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // We will split the table at the third row (inclusive).
        auto row = firstTable->get_Rows()->idx_get(2);
        // Create a new container for the split table.
        auto table = System::DynamicCast<Table>((System::StaticCast<Node>(firstTable))->Clone(false));
        // Insert the container after the original.
        firstTable->get_ParentNode()->InsertAfter(table, firstTable);
        // Add a buffer paragraph to ensure the tables stay apart.
        firstTable->get_ParentNode()->InsertAfter(System::MakeObject<Paragraph>(doc), firstTable);
        System::SharedPtr<Row> currentRow;
        do
        {
            currentRow = firstTable->get_LastRow();
            table->PrependChild(currentRow);
        } while (currentRow != row);

        System::String outputPath = dataDir + GetOutputFilePath(u"JoiningAndSplittingTable.SplitTable.doc");
        // Save the finished document.
        doc->Save(outputPath);
        // ExEnd:SplitTable
        std::cout << "Table splitted successfully into two tables." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void JoiningAndSplittingTable()
{
    std::cout << "JoiningAndSplittingTable example started." << std::endl;
    // ExStart:JoiningAndSplittingTable
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithTables();
    CombineRows(dataDir);
    SplitTable(dataDir);
    // ExEnd:JoiningAndSplittingTable
    std::cout << "JoiningAndSplittingTable example finished." << std::endl << std::endl;
}