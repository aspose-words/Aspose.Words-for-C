#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/TableAlignment.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/TextWrapping.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>

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
                std::cout << System::ObjectExt::Box<VerticalAlignment>(table->get_RelativeVerticalAlignment())->ToString().ToUtf8String() << std::endl;
                std::cout << ".............................." << std::endl;
            }
        }
        // ExEnd:GetFloatingTablePosition
        std::cout << "Get the Table position successfully." << std::endl;
    }

    void SetFloatingTablePosition(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetFloatingTablePosition
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"FloatingTablePosition.docx");

        System::SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        // sets absolute table horizontal position at 10pt.
        table->set_AbsoluteHorizontalDistance(10);

        // sets vertical table position to center of entity specified by Table.VerticalAnchor.
        table->set_RelativeVerticalAlignment(VerticalAlignment::Center);

        // Save the document to disk.
        doc->Save(outputDataDir + u"TablePosition.SetFloatingTablePosition.docx");
        // ExEnd:SetFloatingTablePosition
        std::cout << "Set the Table position successfully." << std::endl;
    }

    void SetRelativeHorizontalOrVerticalPosition(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetRelativeHorizontalOrVerticalPosition
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"FloatingTablePosition.docx");
        System::SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        table->set_HorizontalAnchor(Aspose::Words::Drawing::RelativeHorizontalPosition::Column);
        table->set_VerticalAnchor(Aspose::Words::Drawing::RelativeVerticalPosition::Page);

        // Save the document to disk.
        doc->Save(outputDataDir + u"TablePosition.SetFloatingTablePosition_out.docx");
        // ExEnd:SetRelativeHorizontalOrVerticalPosition
        std::cout << "Set the Table position successfully." << std::endl;
    }
}

void TablePosition()
{
    std::cout << "TablePosition example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithTables();
    System::String outputDataDir = GetOutputDataDir_WorkingWithTables();
    GetTablePosition(inputDataDir);
    GetFloatingTablePosition(inputDataDir);
    SetFloatingTablePosition(inputDataDir, outputDataDir);
    SetRelativeHorizontalOrVerticalPosition(inputDataDir, outputDataDir);
    std::cout << "TablePosition example finished." << std::endl << std::endl;
}