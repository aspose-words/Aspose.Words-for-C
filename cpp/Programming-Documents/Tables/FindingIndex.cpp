#include "stdafx.h"
#include "examples.h"

#include <system/shared_ptr.h>

#include <Model/Document/Document.h>
#include <Model/Tables/Table.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/CellCollection.h>
#include <Model/Tables/Cell.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/Node.h>


using namespace Aspose::Words;

void FindingIndex()
{
    std::cout << "FindingIndex example started." << std::endl;
    // ExStart:FindingIndex
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithTables() + u"Table.SimpleTable.doc";
    auto doc = System::MakeObject<Document>(dataDir);
    // ExStart:RetrieveTableIndex
    // Get the first table in the document.
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    auto allTables = doc->GetChildNodes(Aspose::Words::NodeType::Table, true);
    int32_t tableIndex = allTables->IndexOf(table);
    // ExEnd:RetrieveTableIndex
    std::cout << "Table index is " << tableIndex << std::endl;
    // ExStart:RetrieveRowIndex
    int32_t rowIndex = table->IndexOf(table->get_LastRow());
    // ExEnd:RetrieveRowIndex
    std::cout << "Row index is " << rowIndex << std::endl;
    auto row = table->get_LastRow();
    // ExStart:RetrieveCellIndex
    int32_t cellIndex = row->IndexOf(row->get_Cells()->idx_get(4));
    // ExEnd:RetrieveCellIndex
    std::cout << "Cell index is " << cellIndex << std::endl;
    // ExEnd:FindingIndex
    std::cout << "FindingIndex example finished." << std::endl << std::endl;
}