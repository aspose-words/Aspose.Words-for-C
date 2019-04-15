#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/MailMerge/MailMerge.h>
#include <Model/MailMerge/MappedDataFieldCollection.h>

using namespace Aspose::Words;

namespace
{
    void MappedDataFields()
    {
        // ExStart:MappedDataFields
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        // Shows how to add a mapping when a merge field in a document and a data field in a data source have different names.
        doc->get_MailMerge()->get_MappedDataFields()->Add(u"MyFieldName_InDocument", u"MyFieldName_InDataSource");
        // ExEnd:MappedDataFields
    }

    void DeleteFields()
    {
        // ExStart:DeleteFields
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        // Shows how to delete all merge fields from a document without executing mail merge.
        doc->get_MailMerge()->DeleteFields();
        // ExEnd:DeleteFields
    }
}

void GetFieldNames()
{
    std::cout << "GetFieldNames example started." << std::endl;
    // ExStart:GetFieldNames
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    // Shows how to get names of all merge fields in a document.
    System::ArrayPtr<System::String> fieldNames = doc->get_MailMerge()->GetFieldNames();
    // ExEnd:GetFieldNames
    std::cout << "Document have " << fieldNames->get_Length() << " fields." << std::endl;
    MappedDataFields();
    DeleteFields();
    std::cout << "GetFieldNames example finished." << std::endl << std::endl;
}