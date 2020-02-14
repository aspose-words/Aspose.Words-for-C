#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>


using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void ExtractTableOfContents()
{
    std::cout << "ExtractTableOfContents example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TOC.doc");

    for (System::SharedPtr<Field> field : System::IterateOver(doc->get_Range()->get_Fields()))
    {
        if (field->get_Type() == FieldType::FieldHyperlink)
        {
            System::SharedPtr<FieldHyperlink> hyperlink = System::DynamicCast<FieldHyperlink>(field);
            if (hyperlink->get_SubAddress() != nullptr && hyperlink->get_SubAddress().StartsWith(u"_Toc"))
            {
                System::SharedPtr<Paragraph> tocItem = System::DynamicCast<Paragraph>(field->get_FieldStart()->GetAncestor(NodeType::Paragraph));
                std::cout << System::StaticCast<Node>(tocItem)->ToString(SaveFormat::Text).Trim().ToUtf8String() << std::endl;
                std::cout << "------------------" << std::endl;
                if (tocItem != nullptr)
                {
                    System::SharedPtr<Bookmark> bm = doc->get_Range()->get_Bookmarks()->idx_get(hyperlink->get_SubAddress());
                    // Get the location this TOC Item is pointing to
                    System::SharedPtr<Paragraph> pointer = System::DynamicCast<Paragraph>(bm->get_BookmarkStart()->GetAncestor(NodeType::Paragraph));
                    std::cout << System::StaticCast<Node>(pointer)->ToString(SaveFormat::Text).ToUtf8String() << std::endl;
                }
            }
            // End If
        }
        // End If
    }
    // End Foreach
    std::cout << "ExtractTableOfContents example finished." << std::endl << std::endl;
}