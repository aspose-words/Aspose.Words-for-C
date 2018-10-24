#include "stdafx.h"
#include "examples.h"

#include <system/shared_ptr.h>
#include <system/special_casts.h>
#include <system/object.h>
#include <system/object_ext.h>

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

namespace
{
    void GetTablePosition(System::String const &dataDir)
    {
        // ExStart:GetTablePosition
        auto doc = System::MakeObject<Document>(dataDir + u"Table.Document.doc");
        // Retrieve the first table in the document.
        auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
        if (table->get_TextWrapping() == Aspose::Words::Tables::TextWrapping::Around)
        {
            std::cout << System::ObjectExt::Box<Drawing::HorizontalAlignment>(table->get_RelativeHorizontalAlignment())->ToString().ToUtf8String() << std::endl;
            std::cout << System::ObjectExt::Box<Drawing::VerticalAlignment>(table->get_RelativeVerticalAlignment())->ToString().ToUtf8String() << std::endl;
        }
        else
        {
            std::cout << System::ObjectExt::Box<Aspose::Words::Tables::TableAlignment>(table->get_Alignment())->ToString().ToUtf8String() << std::endl;
        }

        // ExEnd:GetTablePosition
        std::cout << "Get the Table position successfully." << std::endl;
    }

    void GetFloatingTablePosition(System::String const &dataDir)
    {
        // ExStart:GetFloatingTablePosition
        auto doc = System::MakeObject<Document>(dataDir + u"FloatingTablePosition.doc");
        auto table_enumerator = (System::DynamicCastEnumerableTo<System::SharedPtr<Aspose::Words::Tables::Table>>(doc->get_FirstSection()->get_Body()->get_Tables()))->GetEnumerator();
        decltype(table_enumerator->get_Current()) table;
        while (table_enumerator->MoveNext() && (table = table_enumerator->get_Current(), true))
        {
            // If table is floating type then print its positioning properties.
            if (table->get_TextWrapping() == Aspose::Words::Tables::TextWrapping::Around)
            {
                std::cout << System::ObjectExt::Box<Drawing::RelativeHorizontalPosition>(table->get_HorizontalAnchor())->ToString().ToUtf8String() << std::endl;
                std::cout << System::ObjectExt::Box<Drawing::RelativeVerticalPosition>(table->get_VerticalAnchor())->ToString().ToUtf8String() << std::endl;
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
    // ExStart:TablePosition
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithTables();
    GetTablePosition(dataDir);
    GetFloatingTablePosition(dataDir);
    // ExEnd:TablePosition
    std::cout << "TablePosition example finished." << std::endl << std::endl;
}