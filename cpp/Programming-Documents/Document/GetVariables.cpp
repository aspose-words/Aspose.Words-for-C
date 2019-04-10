#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Model/Document/VariableCollection.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void GetVariables()
{
    std::cout << "GetVariables example started." << std::endl;
    // ExStart:GetVariables
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Load the template document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    System::String variables = u"";

    for (System::Collections::Generic::KeyValuePair<System::String, System::String> entry : System::IterateOver(doc->get_Variables()))
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
    std::cout << "Document have following variables " << variables.ToUtf8String() << std::endl;
    std::cout << "GetVariables example finished." << std::endl << std::endl;
}
