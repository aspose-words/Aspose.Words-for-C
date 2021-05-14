#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Bookmark.h>
#include <Aspose.Words.Cpp/BookmarkCollection.h>
#include <Aspose.Words.Cpp/BookmarkEnd.h>
#include <Aspose.Words.Cpp/BookmarkStart.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/ControlChar.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldStart.h>
#include <Aspose.Words.Cpp/ImportFormatMode.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeImporter.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/BookmarksOutlineLevelCollection.h>
#include <Aspose.Words.Cpp/Saving/OutlineOptions.h>
#include <Aspose.Words.Cpp/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <system/array.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/object_ext.h>
#include <system/type_info.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Managment {

class WorkingWithBookmarks : public DocsExamplesBase
{
public:
    void AccessBookmarks()
    {
        //ExStart:AccessBookmarks
        auto doc = MakeObject<Document>(MyDir + u"Bookmarks.docx");

        // By index:
        SharedPtr<Bookmark> bookmark1 = doc->get_Range()->get_Bookmarks()->idx_get(0);
        // By name:
        SharedPtr<Bookmark> bookmark2 = doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark3");
        //ExEnd:AccessBookmarks
    }

    void UpdateBookmarkData()
    {
        //ExStart:UpdateBookmarkData
        auto doc = MakeObject<Document>(MyDir + u"Bookmarks.docx");

        SharedPtr<Bookmark> bookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark1");

        String name = bookmark->get_Name();
        String text = bookmark->get_Text();

        bookmark->set_Name(u"RenamedBookmark");
        bookmark->set_Text(u"This is a new bookmarked text.");
        //ExEnd:UpdateBookmarkData
    }

    void BookmarkTableColumns()
    {
        //ExStart:BookmarkTable
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartTable();

        builder->InsertCell();

        builder->StartBookmark(u"MyBookmark");

        builder->Write(u"This is row 1 cell 1");

        builder->InsertCell();
        builder->Write(u"This is row 1 cell 2");

        builder->EndRow();

        builder->InsertCell();
        builder->Writeln(u"This is row 2 cell 1");

        builder->InsertCell();
        builder->Writeln(u"This is row 2 cell 2");

        builder->EndRow();
        builder->EndTable();

        builder->EndBookmark(u"MyBookmark");
        //ExEnd:BookmarkTable

        //ExStart:BookmarkTableColumns
        for (auto bookmark : System::IterateOver(doc->get_Range()->get_Bookmarks()))
        {
            std::cout << "Bookmark: " << bookmark->get_Name() << (bookmark->get_IsColumn() ? String(u" (Column)") : String(u"")) << std::endl;

            if (bookmark->get_IsColumn())
            {
                auto row = System::DynamicCast_noexcept<Row>(bookmark->get_BookmarkStart()->GetAncestor(NodeType::Row));
                if (row != nullptr && bookmark->get_FirstColumn() < row->get_Cells()->get_Count())
                {
                    std::cout << row->get_Cells()->idx_get(bookmark->get_FirstColumn())->GetText().TrimEnd(MakeArray<char16_t>({ControlChar::CellChar}))
                              << std::endl;
                }
            }
        }
        //ExEnd:BookmarkTableColumns
    }

    void CopyBookmarkedText()
    {
        auto srcDoc = MakeObject<Document>(MyDir + u"Bookmarks.docx");

        // This is the bookmark whose content we want to copy.
        SharedPtr<Bookmark> srcBookmark = srcDoc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark1");

        // We will be adding to this document.
        auto dstDoc = MakeObject<Document>();

        // Let's say we will be appended to the end of the body of the last section.
        SharedPtr<CompositeNode> dstNode = dstDoc->get_LastSection()->get_Body();

        // If you import multiple times without a single context, it will result in many styles created.
        auto importer = MakeObject<NodeImporter>(srcDoc, dstDoc, ImportFormatMode::KeepSourceFormatting);

        AppendBookmarkedText(importer, srcBookmark, dstNode);

        dstDoc->Save(ArtifactsDir + u"WorkingWithBookmarks.CopyBookmarkedText.docx");
    }

    void CreateBookmark()
    {
        //ExStart:CreateBookmark
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartBookmark(u"My Bookmark");
        builder->Writeln(u"Text inside a bookmark.");

        builder->StartBookmark(u"Nested Bookmark");
        builder->Writeln(u"Text inside a NestedBookmark.");
        builder->EndBookmark(u"Nested Bookmark");

        builder->Writeln(u"Text after Nested Bookmark.");
        builder->EndBookmark(u"My Bookmark");

        auto options = MakeObject<PdfSaveOptions>();
        options->get_OutlineOptions()->get_BookmarksOutlineLevels()->Add(u"My Bookmark", 1);
        options->get_OutlineOptions()->get_BookmarksOutlineLevels()->Add(u"Nested Bookmark", 2);

        doc->Save(ArtifactsDir + u"WorkingWithBookmarks.CreateBookmark.pdf", options);
        //ExEnd:CreateBookmark
    }

    void ShowHideBookmarks()
    {
        //ExStart:ShowHideBookmarks
        auto doc = MakeObject<Document>(MyDir + u"Bookmarks.docx");

        ShowHideBookmarkedContent(doc, u"MyBookmark1", false);

        doc->Save(ArtifactsDir + u"WorkingWithBookmarks.ShowHideBookmarks.docx");
        //ExEnd:ShowHideBookmarks
    }

    void ShowHideBookmarkedContent(SharedPtr<Document> doc, String bookmarkName, bool showHide)
    {
        SharedPtr<Bookmark> bm = doc->get_Range()->get_Bookmarks()->idx_get(bookmarkName);

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveToDocumentEnd();

        // {IF "{MERGEFIELD bookmark}" = "true" "" ""}
        SharedPtr<Field> field = builder->InsertField(u"IF \"", nullptr);
        builder->MoveTo(field->get_Start()->get_NextSibling());
        builder->InsertField(String(u"MERGEFIELD ") + bookmarkName + u"", nullptr);
        builder->Write(u"\" = \"true\" ");
        builder->Write(u"\"");
        builder->Write(u"\"");
        builder->Write(u" \"\"");

        SharedPtr<Node> currentNode = field->get_Start();
        bool flag = true;
        while (currentNode != nullptr && flag)
        {
            if (currentNode->get_NodeType() == NodeType::Run)
            {
                if (currentNode->ToString(SaveFormat::Text).Trim() == u"\"")
                {
                    flag = false;
                }
            }

            SharedPtr<Node> nextNode = currentNode->get_NextSibling();

            bm->get_BookmarkStart()->get_ParentNode()->InsertBefore(currentNode, bm->get_BookmarkStart());
            currentNode = nextNode;
        }

        SharedPtr<Node> endNode = bm->get_BookmarkEnd();
        flag = true;
        while (currentNode != nullptr && flag)
        {
            if (currentNode->get_NodeType() == NodeType::FieldEnd)
            {
                flag = false;
            }

            SharedPtr<Node> nextNode = currentNode->get_NextSibling();

            bm->get_BookmarkEnd()->get_ParentNode()->InsertAfter(currentNode, endNode);
            endNode = currentNode;
            currentNode = nextNode;
        }

        doc->get_MailMerge()->Execute(MakeArray<String>({bookmarkName}), MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<bool>(showHide)}));
    }

    void UntangleRowBookmarks()
    {
        auto doc = MakeObject<Document>(MyDir + u"Table column bookmarks.docx");

        // This performs the custom task of putting the row bookmark ends into the same row with the bookmark starts.
        Untangle(doc);

        // Now we can easily delete rows by a bookmark without damaging any other row's bookmarks.
        DeleteRowByBookmark(doc, u"ROW2");

        // This is just to check that the other bookmark was not damaged.
        if (doc->get_Range()->get_Bookmarks()->idx_get(u"ROW1")->get_BookmarkEnd() == nullptr)
        {
            throw System::Exception(u"Wrong, the end of the bookmark was deleted.");
        }

        doc->Save(ArtifactsDir + u"WorkingWithBookmarks.UntangleRowBookmarks.docx");
    }

protected:
    /// <summary>
    /// Copies content of the bookmark and adds it to the end of the specified node.
    /// The destination node can be in a different document.
    /// </summary>
    /// <param name="importer">Maintains the import context.</param>
    /// <param name="srcBookmark">The input bookmark.</param>
    /// <param name="dstNode">Must be a node that can contain paragraphs (such as a Story).</param>
    void AppendBookmarkedText(SharedPtr<NodeImporter> importer, SharedPtr<Bookmark> srcBookmark, SharedPtr<CompositeNode> dstNode)
    {
        // This is the paragraph that contains the beginning of the bookmark.
        auto startPara = System::DynamicCast<Paragraph>(srcBookmark->get_BookmarkStart()->get_ParentNode());

        // This is the paragraph that contains the end of the bookmark.
        auto endPara = System::DynamicCast<Paragraph>(srcBookmark->get_BookmarkEnd()->get_ParentNode());

        if (startPara == nullptr || endPara == nullptr)
        {
            throw System::InvalidOperationException(u"Parent of the bookmark start or end is not a paragraph, cannot handle this scenario yet.");
        }

        // Limit ourselves to a reasonably simple scenario.
        if (startPara->get_ParentNode() != endPara->get_ParentNode())
        {
            throw System::InvalidOperationException(u"Start and end paragraphs have different parents, cannot handle this scenario yet.");
        }

        // We want to copy all paragraphs from the start paragraph up to (and including) the end paragraph,
        // therefore the node at which we stop is one after the end paragraph.
        SharedPtr<Node> endNode = endPara->get_NextSibling();

        for (SharedPtr<Node> curNode = startPara; curNode != endNode; curNode = curNode->get_NextSibling())
        {
            // This creates a copy of the current node and imports it (makes it valid) in the context
            // of the destination document. Importing means adjusting styles and list identifiers correctly.
            SharedPtr<Node> newNode = importer->ImportNode(curNode, true);

            dstNode->AppendChild(newNode);
        }
    }

    void Untangle(SharedPtr<Document> doc)
    {
        for (auto bookmark : System::IterateOver(doc->get_Range()->get_Bookmarks()))
        {
            // Get the parent row of both the bookmark and bookmark end node.
            auto row1 = System::DynamicCast<Row>(bookmark->get_BookmarkStart()->GetAncestor(NodeType::Row));
            auto row2 = System::DynamicCast<Row>(bookmark->get_BookmarkEnd()->GetAncestor(NodeType::Row));

            // If both rows are found okay, and the bookmark start and end are contained in adjacent rows,
            // move the bookmark end node to the end of the last paragraph in the top row's last cell.
            if (row1 != nullptr && row2 != nullptr && row1->get_NextSibling() == row2)
            {
                row1->get_LastCell()->get_LastParagraph()->AppendChild(bookmark->get_BookmarkEnd());
            }
        }
    }

    void DeleteRowByBookmark(SharedPtr<Document> doc, String bookmarkName)
    {
        SharedPtr<Bookmark> bookmark = doc->get_Range()->get_Bookmarks()->idx_get(bookmarkName);
        if (bookmark != nullptr)
        {
            auto row = System::DynamicCast<Row>(bookmark->get_BookmarkStart()->GetAncestor(NodeType::Row));
            if (row != nullptr)
            {
                row->Remove();
            }
        }
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Contents_Managment
