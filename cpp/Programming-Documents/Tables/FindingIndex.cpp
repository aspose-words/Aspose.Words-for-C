#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Tables/Table.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/CellCollection.h>
#include <Model/Tables/Cell.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/Node.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

void FindingIndex()
{
    std::cout << "FindingIndex example started." << std::endl;
    // The path to the documents directory.
    System::String documentPath = GetInputDataDir_WorkingWithTables() + u"Table.SimpleTable.doc";
    System::SharedPtr<Document> doc = System::MakeObject<Document>(documentPath);
    // ExStart:RetrieveTableIndex
    // Get the first table in the document.
    System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
    System::SharedPtr<NodeCollection> allTables = doc->GetChildNodes(NodeType::Table, true);
    int32_t tableIndex = allTables->IndexOf(table);
    // ExEnd:RetrieveTableIndex
    std::cout << "Table index is " << tableIndex << std::endl;
    // ExStart:RetrieveRowIndex
    int32_t rowIndex = table->IndexOf(table->get_LastRow());
    // ExEnd:RetrieveRowIndex
    std::cout << "Row index is " << rowIndex << std::endl;
    System::SharedPtr<Row> row = table->get_LastRow();
    // ExStart:RetrieveCellIndex
    int32_t cellIndex = row->IndexOf(row->get_Cells()->idx_get(4));
    // ExEnd:RetrieveCellIndex
    std::cout << "Cell index is " << cellIndex << std::endl;
    std::cout << "FindingIndex example finished." << std::endl << std::endl;
}