#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/collections/keyvalue_pair.h>
#include <system/collections/ienumerator.h>
#include <Model/Document/VariableCollection.h>
#include <Model/Document/Document.h>
#include <system/object_ext.h>

using namespace Aspose::Words;

void GetVariables()
{
    // ExStart:GetVariables
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Load the template document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    System::String variables = u"";
    
    auto entry_enumerator = doc->get_Variables()->GetEnumerator();
    System::Collections::Generic::KeyValuePair<System::String, System::String> entry;
    while (entry_enumerator->MoveNext() && (entry = entry_enumerator->get_Current(), true))
    {
        System::String name = entry.get_Key();
        System::String value = entry.get_Value();
        if (variables == u"")
        {
            // Do something useful.
            variables = System::String(u"Name: ") + name + u"," + u"Value: " + value;
        }
        else
        {
            variables = variables + u"Name: " + name + u"," + u"Value: " + value;
        }
    }
    // ExEnd:GetVariables
    std::cout << "\nDocument have following variables " << variables.ToUtf8String() << '\n';
}
