#include "stdafx.h"
#include "examples.h"

#include <system/text/regularexpressions/regex.h>
#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/collections/ienumerator.h>
#include <Model/Text/Range.h>
#include <Model/Text/Paragraph.h>
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/Body.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/CompositeNode.h>
#include <Model/Importing/NodeImporter.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/FindReplace/ReplacingArgs.h>
#include <Model/FindReplace/ReplaceAction.h>
#include <Model/FindReplace/IReplacingCallback.h>
#include <Model/FindReplace/FindReplaceOptions.h>
#include <Model/Document/DocumentBase.h>
#include <Model/Document/Document.h>
#include <Model/Bookmarks/BookmarkStart.h>
#include <Model/Bookmarks/BookmarkCollection.h>
#include <Model/Bookmarks/Bookmark.h>
#include <Model/FindReplace/IReplacingCallback.h>
#include <cstdint>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

namespace
{

    // ExStart:InsertDocument
    /// <summary>
    /// Inserts content of the external document after the specified node.
    /// Section breaks and section formatting of the inserted document are ignored.
    /// </summary>
    /// <param name="insertAfterNode">Node in the destination document after which the content
    /// Should be inserted. This node should be a block level node (paragraph or table).</param>
    /// <param name="srcDoc">The document to insert.</param>
    void InsertDocument(System::SharedPtr<Node> insertAfterNode, const System::SharedPtr<Document>& srcDoc)
    {
        // Make sure that the node is either a paragraph or table.
        if (insertAfterNode->get_NodeType() != NodeType::Paragraph && insertAfterNode->get_NodeType() != NodeType::Table)
        {
            throw System::ArgumentException(u"The destination node should be either a paragraph or table.");
        }

        // We will be inserting into the parent of the destination paragraph.
        System::SharedPtr<CompositeNode> dstStory = insertAfterNode->get_ParentNode();

        // This object will be translating styles and lists during the import.
        System::SharedPtr<NodeImporter> importer = System::MakeObject<NodeImporter>(srcDoc, insertAfterNode->get_Document(), Aspose::Words::ImportFormatMode::KeepSourceFormatting);

        // Loop through all sections in the source document.
        auto srcSection_enumerator = srcDoc->get_Sections()->GetEnumerator();
        System::SharedPtr<Section> srcSection;
        while (srcSection_enumerator->MoveNext() && (srcSection = System::DynamicCast<Section>(srcSection_enumerator->get_Current()), true))
        {
            // Loop through all block level nodes (paragraphs and tables) in the body of the section.
            auto srcNode_enumerator = srcSection->get_Body()->GetEnumerator();
            System::SharedPtr<Node> srcNode;
            while (srcNode_enumerator->MoveNext() && (srcNode = srcNode_enumerator->get_Current(), true))
            {
                // Let's skip the node if it is a last empty paragraph in a section.
                if (srcNode->get_NodeType() == Aspose::Words::NodeType::Paragraph)
                {
                    System::SharedPtr<Paragraph> para = System::DynamicCast<Aspose::Words::Paragraph>(srcNode);
                    if (para->get_IsEndOfSection() && !para->get_HasChildNodes())
                    {
                        continue;
                    }
                }

                // This creates a clone of the node, suitable for insertion into the destination document.
                System::SharedPtr<Node> newNode = importer->ImportNode(srcNode, true);

                // Insert new node after the reference node.
                dstStory->InsertAfter(newNode, insertAfterNode);
                insertAfterNode = newNode;
            }
        }
    }
    // ExEnd:InsertDocument

    // ExStart:InsertDocumentAtReplaceHandler
    class InsertDocumentAtReplaceHandler : public IReplacingCallback
    {
        typedef InsertDocumentAtReplaceHandler ThisType;
        typedef Aspose::Words::Replacing::IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);
    public:
        ReplaceAction Replacing(System::SharedPtr<ReplacingArgs> e) override
        {
            System::SharedPtr<Document> subDoc = System::MakeObject<Document>(GetDataDir_WorkingWithDocument() + u"InsertDocument2.doc");

            // Insert a document after the paragraph, containing the match text.
            System::SharedPtr<Paragraph> para = System::DynamicCast<Aspose::Words::Paragraph>(e->get_MatchNode()->get_ParentNode());
            InsertDocument(para, subDoc);

            // Remove the paragraph with the match text.
            para->Remove();

            return Aspose::Words::Replacing::ReplaceAction::Skip;

        }
        
    };
    // ExEnd:InsertDocumentAtReplaceHandler

    void InsertDocumentAtReplace(System::String dataDir)
    {
        // ExStart:InsertDocumentAtReplace
        System::SharedPtr<Document> mainDoc = System::MakeObject<Document>(dataDir + u"InsertDocument1.doc");

        System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(System::MakeObject<InsertDocumentAtReplaceHandler>());

        mainDoc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"\\[MY_DOCUMENT\\]"), u"", options);
        dataDir = dataDir + GetOutputFilePath(u"InsertDoc.InsertDocumentAtReplace.doc");
        mainDoc->Save(dataDir);
        // ExEnd:InsertDocumentAtReplace
        std::cout << "\nDocument inserted successfully at a replace.\nFile saved at " << dataDir.ToUtf8String() << '\n';
    }

    void InsertDocumentAtBookmark(System::String dataDir)
    {
        // ExStart:InsertDocumentAtBookmark
        System::SharedPtr<Document> mainDoc = System::MakeObject<Document>(dataDir + u"InsertDocument1.doc");
        System::SharedPtr<Document> subDoc = System::MakeObject<Document>(dataDir + u"InsertDocument2.doc");

        System::SharedPtr<Bookmark> bookmark = mainDoc->get_Range()->get_Bookmarks()->idx_get(u"insertionPlace");
        InsertDocument(bookmark->get_BookmarkStart()->get_ParentNode(), subDoc);
        dataDir = dataDir + GetOutputFilePath(u"InsertDoc.InsertDocumentAtBookmark.doc");
        mainDoc->Save(dataDir);
        // ExEnd:InsertDocumentAtBookmark
        std::cout << "\nDocument inserted successfully at a bookmark.\nFile saved at " << dataDir.ToUtf8String() << '\n';
    }
}

void InsertDoc()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Invokes the InsertDocument method shown above to insert a document at a bookmark.
    InsertDocumentAtBookmark(dataDir);
    InsertDocumentAtReplace(dataDir);
}
