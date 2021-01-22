#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <system/special_casts.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Importing/NodeImporter.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::MailMerging;

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
        System::SharedPtr<NodeImporter> importer = System::MakeObject<NodeImporter>(srcDoc, insertAfterNode->get_Document(), ImportFormatMode::KeepSourceFormatting);

        // Loop through all sections in the source document.
        for (System::SharedPtr<Section> srcSection : System::IterateOver<System::SharedPtr<Section>>(srcDoc->get_Sections()))
        {
            // Loop through all block level nodes (paragraphs and tables) in the body of the section.
            for (System::SharedPtr<Node> srcNode : System::IterateOver(srcSection->get_Body()))
            {
                // Let's skip the node if it is a last empty paragraph in a section.
                if (srcNode->get_NodeType() == NodeType::Paragraph)
                {
                    System::SharedPtr<Paragraph> para = System::DynamicCast<Paragraph>(srcNode);
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
        typedef IReplacingCallback BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;

    public:
        ReplaceAction Replacing(System::SharedPtr<ReplacingArgs> e) override;
    };

    ReplaceAction InsertDocumentAtReplaceHandler::Replacing(System::SharedPtr<ReplacingArgs> e)
    {
        System::SharedPtr<Document> subDoc = System::MakeObject<Document>(GetInputDataDir_WorkingWithDocument() + u"InsertDocument2.doc");

        // Insert a document after the paragraph, containing the match text.
        System::SharedPtr<Paragraph> para = System::DynamicCast<Paragraph>(e->get_MatchNode()->get_ParentNode());
        InsertDocument(para, subDoc);

        // Remove the paragraph with the match text.
        para->Remove();

        return ReplaceAction::Skip;
    }
    // ExEnd:InsertDocumentAtReplaceHandler

    void InsertDocumentAtReplace(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:InsertDocumentAtReplace
        System::SharedPtr<Document> mainDoc = System::MakeObject<Document>(inputDataDir + u"InsertDocument1.doc");

        System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(System::MakeObject<InsertDocumentAtReplaceHandler>());

        mainDoc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"\\[MY_DOCUMENT\\]"), u"", options);
        System::String outputPath = outputDataDir + u"InsertDoc.InsertDocumentAtReplace.doc";
        mainDoc->Save(outputPath);
        // ExEnd:InsertDocumentAtReplace
        std::cout << "Document inserted successfully at a replace." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertDocumentAtBookmark(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:InsertDocumentAtBookmark
        System::SharedPtr<Document> mainDoc = System::MakeObject<Document>(inputDataDir + u"InsertDocument1.doc");
        System::SharedPtr<Document> subDoc = System::MakeObject<Document>(inputDataDir + u"InsertDocument2.doc");

        System::SharedPtr<Bookmark> bookmark = mainDoc->get_Range()->get_Bookmarks()->idx_get(u"insertionPlace");
        InsertDocument(bookmark->get_BookmarkStart()->get_ParentNode(), subDoc);
        System::String outputPath = outputDataDir + u"InsertDoc.InsertDocumentAtBookmark.doc";
        mainDoc->Save(outputPath);
        // ExEnd:InsertDocumentAtBookmark
        std::cout << "Document inserted successfully at a bookmark." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    // ExStart:InsertDocumentAtMailMergeHandler
    class InsertDocumentAtMailMergeHandler : public IFieldMergingCallback
    {
        typedef InsertDocumentAtMailMergeHandler ThisType;
        typedef IFieldMergingCallback BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;

    public:
        void FieldMerging(System::SharedPtr<FieldMergingArgs> e) override;
        void ImageFieldMerging(System::SharedPtr<ImageFieldMergingArgs> args) override {}
    };

    void InsertDocumentAtMailMergeHandler::FieldMerging(System::SharedPtr<FieldMergingArgs> e)
    {
        if (e->get_DocumentFieldName() == u"Document_1")
        {
            // Use document builder to navigate to the merge field with the specified name.
            System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(e->get_Document());
            builder->MoveToMergeField(e->get_DocumentFieldName());

            // The name of the document to load and insert is stored in the field value.
            System::SharedPtr<Document> subDoc = System::MakeObject<Document>(System::ObjectExt::Unbox<System::String>(e->get_FieldValue()));

            // Insert the document.
            InsertDocument(builder->get_CurrentParagraph(), subDoc);

            // The paragraph that contained the merge field might be empty now and you probably want to delete it.
            if (!builder->get_CurrentParagraph()->get_HasChildNodes())
            {
                builder->get_CurrentParagraph()->Remove();
            }

            // Indicate to the mail merge engine that we have inserted what we wanted.
            e->set_Text(nullptr);
        }
    }
    // ExStart:InsertDocumentAtMailMergeHandler

    void InsertDocumentAtMailMerge(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:InsertDocumentAtMailMerge
        // Open the main document.
        typedef System::SharedPtr<System::Object> TObjectPtr;
        System::SharedPtr<Document> mainDoc = System::MakeObject<Document>(inputDataDir + u"InsertDocument1.doc");

        // Add a handler to MergeField event
        mainDoc->get_MailMerge()->set_FieldMergingCallback(System::MakeObject<InsertDocumentAtMailMergeHandler>());

        // The main document has a merge field in it called "Document_1".
        // The corresponding data for this field contains fully qualified path to the document
        // That should be inserted to this field.
        mainDoc->get_MailMerge()->Execute(System::MakeArray<System::String>({u"Document_1"}), System::StaticCastArray<TObjectPtr>(System::MakeArray<System::String>({inputDataDir + u"InsertDocument2.doc"})));
        System::String outputPath = outputDataDir + u"InsertDoc.InsertDocumentAtMailMerge.doc";
        mainDoc->Save(outputPath);
        // ExEnd:InsertDocumentAtMailMerge
        std::cout << "Document inserted successfully at mail merge." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void InsertDoc()
{
    std::cout << "InsertDoc example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    // Invokes the InsertDocument method shown above to insert a document at a bookmark.
    InsertDocumentAtBookmark(inputDataDir, outputDataDir);
    InsertDocumentAtMailMerge(inputDataDir, outputDataDir);
    InsertDocumentAtReplace(inputDataDir, outputDataDir);
    std::cout << "InsertDoc example finished." << std::endl << std::endl;
}
