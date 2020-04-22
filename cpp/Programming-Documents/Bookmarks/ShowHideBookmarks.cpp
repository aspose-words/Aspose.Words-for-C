#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/array.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>


using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void ShowHideBookmarks()
{
    std::cout << "ShowHideBookmarks example started." << std::endl;
    typedef System::SharedPtr<System::Object> TObjectPtr;

    System::String bookmarkName = u"Bookmark2";

    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithBookmarks();
    System::String outputDataDir = GetOutputDataDir_WorkingWithBookmarks();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Bookmarks.doc");
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    System::SharedPtr<Bookmark> bm = doc->get_Range()->get_Bookmarks()->idx_get(bookmarkName);

    builder->MoveToDocumentEnd();
    // {IF "{MERGEFIELD bookmark}" = "true" "" ""}
    System::SharedPtr<Field> field = builder->InsertField(u"IF \"", nullptr);
    builder->MoveTo(field->get_FieldStart()->get_NextSibling());
    builder->InsertField(System::String(u"MERGEFIELD ") + bookmarkName + u"", nullptr);
    builder->Write(u"\" = \"true\" ");
    builder->Write(u"\"");
    builder->Write(u"\"");
    builder->Write(u" \"\"");

    System::SharedPtr<Node> currentNode = field->get_FieldStart();
    bool flag = true;
    while (currentNode != nullptr && flag)
    {
        if (currentNode->get_NodeType() == Aspose::Words::NodeType::Run)
        {
            if (System::ObjectExt::Equals(currentNode->ToString(Aspose::Words::SaveFormat::Text).Trim(), u"\""))
            {
                flag = false;
            }
        }

        System::SharedPtr<Node> nextNode = currentNode->get_NextSibling();

        bm->get_BookmarkStart()->get_ParentNode()->InsertBefore(currentNode, bm->get_BookmarkStart());
        currentNode = nextNode;
    }

    System::SharedPtr<Node> endNode = bm->get_BookmarkEnd();
    flag = true;
    while (currentNode != nullptr && flag)
    {
        if (currentNode->get_NodeType() == Aspose::Words::NodeType::FieldEnd)
        {
            flag = false;
        }

        System::SharedPtr<Node> nextNode = currentNode->get_NextSibling();

        bm->get_BookmarkEnd()->get_ParentNode()->InsertAfter(currentNode, endNode);
        endNode = currentNode;
        currentNode = nextNode;
    }

    doc->get_MailMerge()->Execute(System::MakeArray<System::String>({bookmarkName}), System::MakeArray<TObjectPtr>({System::ObjectExt::Box<bool>(false)}));

    //MailMerge can be avoided by using the following
    //builder.MoveToMergeField(bookmarkName);
    //builder.Write(showHide ? "true" : "false");

    doc->Save(outputDataDir + u"Updated_Document_out.doc");
    std::cout << "ShowHideBookmarks example finished." << std::endl << std::endl;
}
