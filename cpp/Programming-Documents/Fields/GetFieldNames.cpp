#include "stdafx.h"
#include "examples.h"

#include <system/array.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <Model/Document/Document.h>
#include <Model/MailMerge/MailMerge.h>

using namespace Aspose::Words;

void GetFieldNames()
{
    std::cout << "GetFieldNames example started." << std::endl;
    // ExStart:GetFieldNames
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    // Shows how to get names of all merge fields in a document.
    System::ArrayPtr<System::String> fieldNames = doc->get_MailMerge()->GetFieldNames();
    // ExEnd:GetFieldNames
    std::cout << "Document have " << fieldNames->get_Length() << " fields." << std::endl;
    std::cout << "GetFieldNames example finished." << std::endl << std::endl;
}