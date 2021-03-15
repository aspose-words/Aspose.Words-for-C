#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <system/text/string_builder.h>

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OleFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OleControl.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/Forms2OleControl.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>

using namespace Aspose::Words;

void ReadActiveXControlProperties()
{
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

    auto doc = System::MakeObject<Document>(inputDataDir + u"ActiveX controls.docx");

    System::Text::StringBuilder builder;
    for(const System::SharedPtr<Drawing::Shape> & shape : System::IterateOver<Drawing::Shape>(doc->GetChildNodes(NodeType::Shape, true)))
    {
        if (!shape->get_OleFormat()) 
        {
            break;
        }

        auto oleControl = shape->get_OleFormat()->get_OleControl();
        if (oleControl->get_IsForms2OleControl())
        {
            auto checkbox = System::DynamicCast<Drawing::Ole::Forms2OleControl>(oleControl);
            builder.Append(u"Caption: ");
            builder.AppendLine(checkbox->get_Caption());

            builder.Append(u"Value: ");
            builder.AppendLine(checkbox->get_Value());

            builder.Append(u"Enabled: ");
            builder.AppendLine(System::ObjectExt::ToString(checkbox->get_Enabled()));

            builder.Append(u"Type: ");
            builder.AppendLine(System::ObjectExt::ToString(checkbox->get_Type()));

            if (checkbox->get_ChildNodes())
            {
                builder.AppendLine(u"ChildNodes: true");
            }
        }

    }

    builder.Append(u"Total ActiveX Controls found: ");
    builder.AppendFormat(u"{0}\n", doc->GetChildNodes(NodeType::Shape, true)->get_Count());

    std::cout << builder.ToString().ToUtf8String() << '\n';
}
