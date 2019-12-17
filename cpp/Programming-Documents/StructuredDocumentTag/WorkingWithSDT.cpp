#include "stdafx.h"
#include "examples.h"

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
#include <Model/Styles/StyleCollection.h>
#include <Model/Styles/StyleIdentifier.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;

namespace
{
    void SetContentControlColor(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetContentControlColor

        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"input.docx");
        System::SharedPtr<StructuredDocumentTag> sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        sdt->set_Color(System::Drawing::Color::get_Red());

        System::String outputPath = outputDataDir + u"WorkingWithSDT.SetContentControlColor.docx";

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetContentControlColor
        std::cout << "Set the color of content control successfully." << std::endl;
    }

    void ClearContentsControl(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:ClearContentsControl
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"input.docx");
        System::SharedPtr<StructuredDocumentTag> sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        sdt->Clear();

        System::String outputPath = outputDataDir + u"WorkingWithSDT.ClearContentsControl.doc";

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:ClearContentsControl
        std::cout << "Clear the contents of content control successfully." << std::endl;
    }

    void BindSDTtoCustomXmlPart(System::String const &outputDataDir)
    {
        // ExStart:BindSDTtoCustomXmlPart
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(System::Guid::NewGuid().ToString(u"B"), u"<root><text>Hello, World!</text></root>");

        System::SharedPtr<StructuredDocumentTag> sdt = System::MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Block);
        doc->get_FirstSection()->get_Body()->AppendChild(sdt);
        sdt->get_XmlMapping()->SetMapping(xmlPart, u"/root[1]/text[1]", u"");

        System::String outputPath = outputDataDir + u"WorkingWithSDT.BindSDTtoCustomXmlPart.doc";

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:BindSDTtoCustomXmlPart
        std::cout << "Creation of an XML part and binding a content control to it successfully." << std::endl;
    }

    void SetContentControlStyle(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetContentControlStyle
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"input.docx");
        System::SharedPtr<StructuredDocumentTag> sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        System::SharedPtr<Style> style = doc->get_Styles()->idx_get(StyleIdentifier::Quote);
        sdt->set_Style(style);

        System::String outputPath = outputDataDir + u"WorkingWithSDT.SetContentControlStyle.docx";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetContentControlStyle
        std::cout << "Set the style of content control successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkingWithSDT()
{
    std::cout << "WorkingWithSDT example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithStructuredDocumentTag();
    System::String outputDataDir = GetOutputDataDir_WorkingWithStructuredDocumentTag();
    BindSDTtoCustomXmlPart(outputDataDir);
    ClearContentsControl(inputDataDir, outputDataDir);
    SetContentControlColor(inputDataDir, outputDataDir);
    SetContentControlStyle(inputDataDir, outputDataDir);
    std::cout << "WorkingWithSDT example finished." << std::endl << std::endl;
}