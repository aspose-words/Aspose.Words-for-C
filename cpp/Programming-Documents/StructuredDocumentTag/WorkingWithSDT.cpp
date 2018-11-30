#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/special_casts.h>
#include <Model/Document/Document.h>
#include <Model/Markup/Sdt/SdtType.h>
#include <Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Model/Markup/Sdt/XmlMapping.h>
#include <Model/Markup/CustomXmlPart.h>
#include <Model/Markup/CustomXmlPartCollection.h>
#include <Model/Markup/MarkupLevel.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/Section.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;

namespace
{
    void SetContentControlColor(System::String const &dataDir)
    {
        // ExStart:SetContentControlColor

        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"input.docx");
        System::SharedPtr<StructuredDocumentTag> sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        sdt->set_Color(System::Drawing::Color::get_Red());

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithSDT.SetContentControlColor.docx");

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetContentControlColor
        std::cout << "Set the color of content control successfully." << std::endl;
    }

    void ClearContentsControl(System::String const &dataDir)
    {
        // ExStart:ClearContentsControl
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"input.docx");
        System::SharedPtr<StructuredDocumentTag> sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        sdt->Clear();

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithSDT.ClearContentsControl.doc");

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:ClearContentsControl
        std::cout << "Clear the contents of content control successfully." << std::endl;
    }

    void BindSDTtoCustomXmlPart(System::String const &dataDir)
    {
        // ExStart:BindSDTtoCustomXmlPart
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(System::Guid::NewGuid().ToString(u"B"), u"<root><text>Hello, World!</text></root>");

        System::SharedPtr<StructuredDocumentTag> sdt = System::MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Block);
        doc->get_FirstSection()->get_Body()->AppendChild(sdt);
        sdt->get_XmlMapping()->SetMapping(xmlPart, u"/root[1]/text[1]", u"");

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithSDT.BindSDTtoCustomXmlPart.doc");

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:BindSDTtoCustomXmlPart
        std::cout << "Creation of an XML part and binding a content control to it successfully." << std::endl;
    }
}

void WorkingWithSDT()
{
    std::cout << "WorkingWithSDT example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithStructuredDocumentTag();
    BindSDTtoCustomXmlPart(dataDir);
    ClearContentsControl(dataDir);
    SetContentControlColor(dataDir);
    std::cout << "WorkingWithSDT example finished." << std::endl << std::endl;
}