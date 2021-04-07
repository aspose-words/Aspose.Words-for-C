#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/Forms2OleControl.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/Forms2OleControlCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/Forms2OleControlType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OleControl.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OleFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OlePackage.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <system/array.h>
#include <system/details/dispose_guard.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/io/file.h>
#include <system/io/memory_stream.h>
#include <system/io/stream.h>
#include <system/object_ext.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Ole;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithOleObjectsAndActiveX : public DocsExamplesBase
{
public:
    void InsertOleObject()
    {
        //ExStart:DocumentBuilderInsertOleObject
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertOleObject(u"http://www.aspose.com", u"htmlfile", true, true, nullptr);

        doc->Save(ArtifactsDir + u"WorkingWithOleObjectsAndActiveX.InsertOleObject.docx");
        //ExEnd:DocumentBuilderInsertOleObject
    }

    void InsertOleObjectWithOlePackage()
    {
        //ExStart:InsertOleObjectwithOlePackage
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        ArrayPtr<uint8_t> bs = System::IO::File::ReadAllBytes(MyDir + u"Zip file.zip");
        {
            SharedPtr<System::IO::Stream> stream = MakeObject<System::IO::MemoryStream>(bs);
            SharedPtr<Shape> shape = builder->InsertOleObject(stream, u"Package", true, nullptr);
            SharedPtr<OlePackage> olePackage = shape->get_OleFormat()->get_OlePackage();
            olePackage->set_FileName(u"filename.zip");
            olePackage->set_DisplayName(u"displayname.zip");

            doc->Save(ArtifactsDir + u"WorkingWithOleObjectsAndActiveX.InsertOleObjectWithOlePackage.docx");
        }
        //ExEnd:InsertOleObjectwithOlePackage

        //ExStart:GetAccessToOLEObjectRawData
        auto oleShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        ArrayPtr<uint8_t> oleRawData = oleShape->get_OleFormat()->GetRawData();
        //ExEnd:GetAccessToOLEObjectRawData
    }

    void InsertOleObjectAsIcon()
    {
        //ExStart:InsertOLEObjectAsIcon
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertOleObjectAsIcon(MyDir + u"Presentation.pptx", false, ImagesDir + u"Logo icon.ico", u"My embedded file");

        doc->Save(ArtifactsDir + u"WorkingWithOleObjectsAndActiveX.InsertOleObjectAsIcon.docx");
        //ExEnd:InsertOLEObjectAsIcon
    }

    void InsertOleObjectAsIconUsingStream()
    {
        //ExStart:InsertOLEObjectAsIconUsingStream
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        {
            auto stream = MakeObject<System::IO::MemoryStream>(System::IO::File::ReadAllBytes(MyDir + u"Presentation.pptx"));
            builder->InsertOleObjectAsIcon(stream, u"Package", ImagesDir + u"Logo icon.ico", u"My embedded file");
        }

        doc->Save(ArtifactsDir + u"WorkingWithOleObjectsAndActiveX.InsertOleObjectAsIconUsingStream.docx");
        //ExEnd:InsertOLEObjectAsIconUsingStream
    }

    void ReadActiveXControlProperties()
    {
        auto doc = MakeObject<Document>(MyDir + u"ActiveX controls.docx");

        String properties = u"";
        for (auto shape : System::IterateOver<Shape>(doc->GetChildNodes(NodeType::Shape, true)))
        {
            if (shape->get_OleFormat() == nullptr)
            {
                break;
            }

            SharedPtr<OleControl> oleControl = shape->get_OleFormat()->get_OleControl();
            if (oleControl->get_IsForms2OleControl())
            {
                auto checkBox = System::DynamicCast<Forms2OleControl>(oleControl);
                properties = properties + u"\nCaption: " + checkBox->get_Caption();
                properties = properties + u"\nValue: " + checkBox->get_Value();
                properties = properties + u"\nEnabled: " + checkBox->get_Enabled();
                properties = properties + u"\nType: " + System::ObjectExt::ToString(checkBox->get_Type());
                if (checkBox->get_ChildNodes() != nullptr)
                {
                    properties = properties + u"\nChildNodes: " + checkBox->get_ChildNodes();
                }

                properties += u"\n";
            }
        }

        properties = properties + u"\nTotal ActiveX Controls found: " + doc->GetChildNodes(NodeType::Shape, true)->get_Count();
        std::cout << (String(u"\n") + properties) << std::endl;
    }
};

}} // namespace DocsExamples::Programming_with_Documents
