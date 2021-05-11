#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/BookmarkEnd.h>
#include <Aspose.Words.Cpp/BookmarkStart.h>
#include <Aspose.Words.Cpp/ControlChar.h>
#include <Aspose.Words.Cpp/ConvertUtil.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Properties/CustomDocumentProperties.h>
#include <Aspose.Words.Cpp/Properties/DocumentProperty.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/VariableCollection.h>
#include <system/collections/keyvalue_pair.h>
#include <system/date_time.h>
#include <system/enumerator_adapter.h>
#include <system/object_ext.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Properties;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document {

class DocumentPropertiesAndVariables : public DocsExamplesBase
{
public:
    void GetVariables()
    {
        //ExStart:GetVariables
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        String variables = u"";
        for (auto entry : doc->get_Variables())
        {
            String name = entry.get_Key();
            String value = entry.get_Value();
            if (variables == u"")
            {
                variables = String(u"Name: ") + name + u"," + u"Value: {1}" + value;
            }
            else
            {
                variables = variables + u"Name: " + name + u"," + u"Value: {1}" + value;
            }
        }
        //ExEnd:GetVariables

        std::cout << (String(u"\nDocument have following variables ") + variables) << std::endl;
    }

    void EnumerateProperties()
    {
        //ExStart:EnumerateProperties
        auto doc = MakeObject<Document>(MyDir + u"Properties.docx");

        std::cout << "1. Document name: " << doc->get_OriginalFileName() << std::endl;
        std::cout << "2. Built-in Properties" << std::endl;

        for (auto prop : System::IterateOver(doc->get_BuiltInDocumentProperties()))
        {
            std::cout << prop->get_Name() << " : " << prop->get_Value() << std::endl;
        }

        std::cout << "3. Custom Properties" << std::endl;

        for (auto prop : System::IterateOver(doc->get_CustomDocumentProperties()))
        {
            std::cout << prop->get_Name() << " : " << prop->get_Value() << std::endl;
        }
        //ExEnd:EnumerateProperties
    }

    void AddCustomDocumentProperties()
    {
        //ExStart:AddCustomDocumentProperties
        auto doc = MakeObject<Document>(MyDir + u"Properties.docx");

        SharedPtr<CustomDocumentProperties> customDocumentProperties = doc->get_CustomDocumentProperties();

        if (customDocumentProperties->idx_get(u"Authorized") != nullptr)
        {
            return;
        }

        customDocumentProperties->Add(u"Authorized", true);
        customDocumentProperties->Add(u"Authorized By", String(u"John Smith"));
        customDocumentProperties->Add(u"Authorized Date", System::DateTime::get_Today());
        customDocumentProperties->Add(u"Authorized Revision", doc->get_BuiltInDocumentProperties()->get_RevisionNumber());
        customDocumentProperties->Add(u"Authorized Amount", 123.45);
        //ExEnd:AddCustomDocumentProperties
    }

    void RemoveCustomDocumentProperties()
    {
        //ExStart:CustomRemove
        auto doc = MakeObject<Document>(MyDir + u"Properties.docx");
        doc->get_CustomDocumentProperties()->Remove(u"Authorized Date");
        //ExEnd:CustomRemove
    }

    void RemovePersonalInformation()
    {
        //ExStart:RemovePersonalInformation
        auto doc = MakeObject<Document>(MyDir + u"Properties.docx");
        doc->set_RemovePersonalInformation(true);

        doc->Save(ArtifactsDir + u"DocumentPropertiesAndVariables.RemovePersonalInformation.docx");
        //ExEnd:RemovePersonalInformation
    }

    void ConfiguringLinkToContent()
    {
        //ExStart:ConfiguringLinkToContent
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartBookmark(u"MyBookmark");
        builder->Writeln(u"Text inside a bookmark.");
        builder->EndBookmark(u"MyBookmark");

        // Retrieve a list of all custom document properties from the file.
        SharedPtr<CustomDocumentProperties> customProperties = doc->get_CustomDocumentProperties();
        // Add linked to content property.
        SharedPtr<DocumentProperty> customProperty = customProperties->AddLinkToContent(u"Bookmark", u"MyBookmark");
        customProperty = customProperties->idx_get(u"Bookmark");

        bool isLinkedToContent = customProperty->get_IsLinkToContent();

        String linkSource = customProperty->get_LinkSource();

        String customPropertyValue = System::ObjectExt::ToString(customProperty->get_Value());
        //ExEnd:ConfiguringLinkToContent
    }

    void ConvertBetweenMeasurementUnits()
    {
        //ExStart:ConvertBetweenMeasurementUnits
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<PageSetup> pageSetup = builder->get_PageSetup();
        pageSetup->set_TopMargin(ConvertUtil::InchToPoint(1.0));
        pageSetup->set_BottomMargin(ConvertUtil::InchToPoint(1.0));
        pageSetup->set_LeftMargin(ConvertUtil::InchToPoint(1.5));
        pageSetup->set_RightMargin(ConvertUtil::InchToPoint(1.5));
        pageSetup->set_HeaderDistance(ConvertUtil::InchToPoint(0.2));
        pageSetup->set_FooterDistance(ConvertUtil::InchToPoint(0.2));
        //ExEnd:ConvertBetweenMeasurementUnits
    }

    void UseControlCharacters()
    {
        //ExStart:UseControlCharacters
        const String text = u"test\r";
        // Replace "\r" control character with "\r\n".
        String replace = text.Replace(ControlChar::Cr(), ControlChar::CrLf());
        //ExEnd:UseControlCharacters
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document
