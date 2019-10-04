#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Properties/DocumentProperty.h>
#include <Aspose.Words.Cpp/Model/Properties/CustomDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Properties;

namespace
{
    void EnumerateProperties(const System::String& inputDataDir)
    {
        // ExStart:EnumerateProperties
        System::String fileName = inputDataDir + u"Properties.doc";
        System::SharedPtr<Document> doc = System::MakeObject<Document>(fileName);
        std::cout << "1. Document name: " << fileName.ToUtf8String() << std::endl;

        std::cout << "2. Built-in Properties" << std::endl;
        for (System::SharedPtr<DocumentProperty> prop : System::IterateOver(doc->get_BuiltInDocumentProperties()))
        {
            std::cout << prop->get_Name().ToUtf8String() << " : " << prop->get_Value()->ToString().ToUtf8String() << std::endl;
        }

        std::cout << "3. Custom Properties" << std::endl;
        for (System::SharedPtr<DocumentProperty> prop : System::IterateOver(doc->get_CustomDocumentProperties()))
        {
            std::cout << prop->get_Name().ToUtf8String() << " : " << prop->get_Value()->ToString().ToUtf8String() << std::endl;
        }
        // ExEnd:EnumerateProperties
    }

    void CustomAdd(const System::String& inputDataDir)
    {
        // ExStart:CustomAdd
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Properties.doc");
        System::SharedPtr<CustomDocumentProperties> props = doc->get_CustomDocumentProperties();
        if (props->idx_get(u"Authorized") == nullptr)
        {
            props->Add(u"Authorized", true);
            props->Add(u"Authorized By", System::String(u"John Smith"));
            props->Add(u"Authorized Date", System::DateTime::get_Today());
            props->Add(u"Authorized Revision", doc->get_BuiltInDocumentProperties()->get_RevisionNumber());
            props->Add(u"Authorized Amount", 123.45);
        }
        // ExEnd:CustomAdd
    }

    void CustomRemove(const System::String& inputDataDir)
    {
        // ExStart:CustomRemove
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Properties.doc");
        doc->get_CustomDocumentProperties()->Remove(u"Authorized Date");
        // ExEnd:CustomRemove
    }

    void RemovePersonalInformation(const System::String &inputDataDir, const System::String &outputDataDir)
    {
        // ExStart:RemovePersonalInformation
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Properties.doc");
        doc->set_RemovePersonalInformation(true);
        System::String outputPath = outputDataDir + u"DocProperties.RemovePersonalInformation.docx";
        doc->Save(outputPath);
        // ExEnd:RemovePersonalInformation
        std::cout << "Personal information has removed from document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void DocProperties()
{
    std::cout << "DocProperties example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    // Enumerates through all built-in and custom properties in a document.
    EnumerateProperties(inputDataDir);
    // Checks if a custom property with a given name exists in a document and adds few more custom document properties.
    CustomAdd(inputDataDir);
    // Removes a custom document property.
    CustomRemove(inputDataDir);
    RemovePersonalInformation(inputDataDir, outputDataDir);
    std::cout << "DocProperties example finished." << std::endl << std::endl;
}
