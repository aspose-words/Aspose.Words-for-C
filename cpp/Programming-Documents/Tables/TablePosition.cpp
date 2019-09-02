#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include "Model/Document/Document.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/TableAlignment.h>
#include <Model/Tables/TableCollection.h>
#include <Model/Tables/TextWrapping.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/Body.h>
#include <Model/Drawing/HorizontalAlignment.h>
#include <Model/Drawing/VerticalAlignment.h>
#include <Model/Drawing/RelativeVerticalPosition.h>
#include <Model/Drawing/RelativeHorizontalPosition.h>
#include <Model/Nodes/NodeType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Tables;

namespace
{
    void GetTablePosition(System::String const &inputDataDir)
    {
        // ExStart:GetTablePosition
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.Document.doc");
        // Retrieve the first table in the document.
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        if (table->get_TextWrapping() == TextWrapping::Around)
        {
            std::cout << System::ObjectExt::Box<HorizontalAlignment>(table->get_RelativeHorizontalAlignment())->ToString().ToUtf8String() << std::endl;
            std::cout << System::ObjectExt::Box<VerticalAlignment>(table->get_RelativeVerticalAlignment())->ToString().ToUtf8String() << std::endl;
        }
        else
        {
            std::cout << System::ObjectExt::Box<TableAlignment>(table->get_Alignment())->ToString().ToUtf8String() << std::endl;
        }

        // ExEnd:GetTablePosition
        std::cout << "Get the Table position successfully." << std::endl;
    }

    void GetFloatingTablePosition(System::String const &inputDataDir)
    {
        // ExStart:GetFloatingTablePosition
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"FloatingTablePosition.docx");
        for (System::SharedPtr<Table> table : System::IterateOver<System::SharedPtr<Table>>(doc->get_FirstSection()->get_Body()->get_Tables()))
        {
            // If table is floating type then print its positioning properties.
            if (table->get_TextWrapping() == TextWrapping::Around)
            {
                std::cout << System::ObjectExt::Box<RelativeHorizontalPosition>(table->get_HorizontalAnchor())->ToString().ToUtf8String() << std::endl;
                std::cout << System::ObjectExt::Box<RelativeVerticalPosition>(table->get_VerticalAnchor())->ToString().ToUtf8String() << std::endl;
                std::cout << table->get_AbsoluteHorizontalDistance() << std::endl;
                std::cout << table->get_AbsoluteVerticalDistance() << std::endl;
                std::cout << table->get_AllowOverlap() << std::endl;
                std::cout << ".............................." << std::endl;
            }
        }
        // ExEnd:GetFloatingTablePosition
        std::cout << "Get the Table position successfully." << std::endl;
    }
}

void TablePosition()
{
    std::cout << "TablePosition example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithTables();
    GetTablePosition(inputDataDir);
    GetFloatingTablePosition(inputDataDir);
    std::cout << "TablePosition example finished." << std::endl << std::endl;
}