#include "stdafx.h"
#include "examples.h"

#include <system/shared_ptr.h>

#include "Model/Document/Document.h"
#include <Model/Fields/FieldCollection.h>
#include <Model/Fields/FieldType.h>
#include <Model/Fields/Fields/LinksAndReferences/FieldHyperlink.h>
#include <Model/Text/Range.h>

using namespace Aspose::Words;

void ReplaceHyperlinks()
{
    // ExStart:ReplaceHyperlinks
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithHyperlink();
    System::String NewUrl = u"http://www.aspose.com";
    System::String NewName = u"Aspose - The .NET & Java Component Publisher";
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"ReplaceHyperlinks.doc");
    // Hyperlinks in a Word documents are fields.
    auto field_enumerator = (doc->get_Range()->get_Fields())->GetEnumerator();
    decltype(field_enumerator->get_Current()) field;
    while (field_enumerator->MoveNext() && (field = field_enumerator->get_Current(), true))
    {
        if (field->get_Type() == Aspose::Words::Fields::FieldType::FieldHyperlink)
        {
            auto hyperlink = System::DynamicCast<Aspose::Words::Fields::FieldHyperlink>(field);
            // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
            if (hyperlink->get_SubAddress() != nullptr)
            {
                continue;
            }
            hyperlink->set_Address(NewUrl);
            hyperlink->set_Result(NewName);
        }
    }
    System::String outputPath = dataDir + GetOutputFilePath(u"ReplaceHyperlinks.doc");
    doc->Save(outputPath);
    // ExEnd:ReplaceHyperlinks
    std::cout << "\nHyperlinks replaced successfully.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
}