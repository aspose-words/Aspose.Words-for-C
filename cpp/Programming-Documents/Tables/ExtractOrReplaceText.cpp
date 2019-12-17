#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceDirection.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Tables;

namespace
{
    void ExtractPrintText(System::String const &documentPath)
    {
        // ExStart:ExtractText
        System::SharedPtr<Document> doc = System::MakeObject<Document>(documentPath);
        // Get the first table in the document.
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // The range text will include control characters such as "\a" for a cell.
        // You can call ToString and pass SaveFormat.Text on the desired node to find the plain text content.
        // Print the plain text range of the table to the screen.
        std::cout << "Contents of the table: " << std::endl << table->get_Range()->get_Text().ToUtf8String() << std::endl;
        // ExEnd:ExtractText
        // ExStart:PrintTextRangeOFRowAndTable
        // Print the contents of the second row to the screen.
        std::cout << "Contents of the row: " << std::endl << table->get_Rows()->idx_get(1)->get_Range()->get_Text().ToUtf8String() << std::endl;
        // Print the contents of the last cell in the table to the screen.
        std::cout << "Contents of the cell: " << std::endl << table->get_LastRow()->get_LastCell()->get_Range()->get_Text().ToUtf8String() << std::endl;
        // ExEnd:PrintTextRangeOFRowAndTable
    }

    void ReplaceText(System::String const &documentPath, System::String const &outputDataDir)
    {
        // ExStart:ReplaceText
        System::SharedPtr<Document> doc = System::MakeObject<Document>(documentPath);
        // Get the first table in the document.
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // Replace any instances of our string in the entire table.
        table->get_Range()->Replace(u"Carrots", u"Eggs", System::MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));
        // Replace any instances of our string in the last cell of the table only.
        table->get_LastRow()->get_LastCell()->get_Range()->Replace(u"50", u"20", System::MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));
        System::String outputPath = outputDataDir + u"ExtractOrReplaceText.ReplaceText.doc";
        doc->Save(outputPath);
        // ExEnd:ReplaceText
        std::cout << "Text replaced successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void ExtractOrReplaceText()
{
    std::cout << "ExtractOrReplaceText example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithTables();
    System::String outputDataDir = GetOutputDataDir_WorkingWithTables();
    System::String documentPath = inputDataDir + u"Table.SimpleTable.doc";
    ExtractPrintText(documentPath);
    ReplaceText(documentPath, outputDataDir);
    std::cout << "ExtractOrReplaceText example finished." << std::endl << std::endl;
}