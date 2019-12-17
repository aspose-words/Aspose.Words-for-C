#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include "Model/Document/Document.h"
#include <Model/Fields/FieldCollection.h>
#include <Model/Fields/FieldType.h>
#include <Model/Fields/Fields/LinksAndReferences/FieldHyperlink.h>
#include <Model/Text/Range.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void ReplaceHyperlinks()
{
    std::cout << "ReplaceHyperlinks example started." << std::endl;
    // ExStart:ReplaceHyperlinks
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithHyperlink();
    System::String outputDataDir = GetOutputDataDir_WorkingWithHyperlink();
    System::String NewUrl = u"http://www.aspose.com";
    System::String NewName = u"Aspose - The .NET & Java Component Publisher";
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"ReplaceHyperlinks.doc");
    // Hyperlinks in a Word documents are fields.
    for (System::SharedPtr<Field> field : System::IterateOver(doc->get_Range()->get_Fields()))
    {
        if (field->get_Type() == FieldType::FieldHyperlink)
        {
            System::SharedPtr<FieldHyperlink> hyperlink = System::DynamicCast<FieldHyperlink>(field);
            // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
            if (hyperlink->get_SubAddress() != nullptr)
            {
                continue;
            }
            hyperlink->set_Address(NewUrl);
            hyperlink->set_Result(NewName);
        }
    }
    System::String outputPath = outputDataDir + u"ReplaceHyperlinks.doc";
    doc->Save(outputPath);
    // ExEnd:ReplaceHyperlinks
    std::cout << "Hyperlinks replaced successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ReplaceHyperlinks example finished." << std::endl << std::endl;
}