#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Bookmark.h>
#include <Aspose.Words.Cpp/BookmarkCollection.h>
#include <Aspose.Words.Cpp/BookmarkStart.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/ImportFormatMode.h>
#include <Aspose.Words.Cpp/MailMerging/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/MailMerging/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/MailMerging/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeImporter.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Replacing/FindReplaceDirection.h>
#include <Aspose.Words.Cpp/Replacing/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Replacing/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Replacing/ReplaceAction.h>
#include <Aspose.Words.Cpp/Replacing/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/io/memory_stream.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/text/regularexpressions/regex.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
using namespace Aspose::Words::Replacing;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document {

class CloneAndCombineDocuments : public DocsExamplesBase
{
public:
    void CloningDocument()
    {
        //ExStart:CloningDocument
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<Document> clone = doc->Clone();
        clone->Save(ArtifactsDir + u"CloneAndCombineDocuments.CloningDocument.docx");
        //ExEnd:CloningDocument
    }

    void InsertDocumentAtReplace()
    {
        //ExStart:InsertDocumentAtReplace
        //GistId:db2dfc4150d7c714bcac3782ae241d03
        auto mainDoc = MakeObject<Document>(MyDir + u"Document insertion 1.docx");

        // Set find and replace options.
        auto options = MakeObject<FindReplaceOptions>();
        options->set_Direction(FindReplaceDirection::Backward);
        options->set_ReplacingCallback(MakeObject<CloneAndCombineDocuments::InsertDocumentAtReplaceHandler>());

        // Call the replace method.
        mainDoc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"\\[MY_DOCUMENT\\]"), u"", options);
        mainDoc->Save(ArtifactsDir + u"CloneAndCombineDocuments.InsertDocumentAtReplace.docx");
        //ExEnd:InsertDocumentAtReplace
    }

    void InsertDocumentAtBookmark()
    {
        //ExStart:InsertDocumentAtBookmark
        //GistId:db2dfc4150d7c714bcac3782ae241d03
        auto mainDoc = MakeObject<Document>(MyDir + u"Document insertion 1.docx");
        auto subDoc = MakeObject<Document>(MyDir + u"Document insertion 2.docx");

        SharedPtr<Bookmark> bookmark = mainDoc->get_Range()->get_Bookmarks()->idx_get(u"insertionPlace");
        InsertDocument(bookmark->get_BookmarkStart()->get_ParentNode(), subDoc);

        mainDoc->Save(ArtifactsDir + u"CloneAndCombineDocuments.InsertDocumentAtBookmark.docx");
        //ExEnd:InsertDocumentAtBookmark
    }

    void InsertDocumentAtMailMerge()
    {
        //ExStart:InsertDocumentAtMailMerge
        //GistId:db2dfc4150d7c714bcac3782ae241d03
        auto mainDoc = MakeObject<Document>(MyDir + u"Document insertion 1.docx");

        mainDoc->get_MailMerge()->set_FieldMergingCallback(MakeObject<CloneAndCombineDocuments::InsertDocumentAtMailMergeHandler>());
        // The main document has a merge field in it called "Document_1".
        // The corresponding data for this field contains a fully qualified path to the document.
        // That should be inserted to this field.
        mainDoc->get_MailMerge()->Execute(MakeArray<String>({u"Document_1"}),
                                          MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(MyDir + u"Document insertion 2.docx")}));

        mainDoc->Save(ArtifactsDir + u"CloneAndCombineDocuments.InsertDocumentAtMailMerge.doc");
        //ExEnd:InsertDocumentAtMailMerge
    }

    //ExStart:InsertDocumentAsNodes
    //GistId:db2dfc4150d7c714bcac3782ae241d03
    /// <summary>
    /// Inserts content of the external document after the specified node.
    /// Section breaks and section formatting of the inserted document are ignored.
    /// </summary>
    /// <param name="insertionDestination">Node in the destination document after which the content
    /// Should be inserted. This node should be a block level node (paragraph or table).</param>
    /// <param name="docToInsert">The document to insert.</param>    
    static void InsertDocument(SharedPtr<Node> insertionDestination, SharedPtr<Document> docToInsert)
    {
        if (insertionDestination->get_NodeType() == NodeType::Paragraph || insertionDestination->get_NodeType() == NodeType::Table)
        {
            SharedPtr<CompositeNode> destinationParent = insertionDestination->get_ParentNode();

            auto importer = MakeObject<NodeImporter>(docToInsert, insertionDestination->get_Document(), ImportFormatMode::KeepSourceFormatting);

            // Loop through all block-level nodes in the section's body,
            // then clone and insert every node that is not the last empty paragraph of a section.
            for (const auto& srcSection : System::IterateOver(docToInsert->get_Sections()->LINQ_OfType<SharedPtr<Section>>()))
            {
                for (const auto& srcNode : System::IterateOver(srcSection->get_Body()))
                {
                    if (srcNode->get_NodeType() == NodeType::Paragraph)
                    {
                        auto para = System::ExplicitCast<Paragraph>(srcNode);
                        if (para->get_IsEndOfSection() && !para->get_HasChildNodes())
                        {
                            continue;
                        }
                    }

                    SharedPtr<Node> newNode = importer->ImportNode(srcNode, true);

                    destinationParent->InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
        else
        {
            throw System::ArgumentException(u"The destination node should be either a paragraph or table.");
        }
    }
    //ExEnd:InsertDocumentAsNodes

    //ExStart:InsertDocumentWithSectionFormatting
    /// <summary>
    /// Inserts content of the external document after the specified node.
    /// </summary>
    /// <param name="insertAfterNode">Node in the destination document after which the content
    /// Should be inserted. This node should be a block level node (paragraph or table).</param>
    /// <param name="srcDoc">The document to insert.</param>
    void InsertDocumentWithSectionFormatting(SharedPtr<Node> insertAfterNode, SharedPtr<Document> srcDoc)
    {
        if (insertAfterNode->get_NodeType() != NodeType::Paragraph && insertAfterNode->get_NodeType() != NodeType::Table)
        {
            throw System::ArgumentException(u"The destination node should be either a paragraph or table.");
        }

        auto dstDoc = System::ExplicitCast<Document>(insertAfterNode->get_Document());
        // To retain section formatting, split the current section into two at the marker node and then import the content
        // from srcDoc as whole sections. The section of the node to which the insert marker node belongs.
        auto currentSection = System::ExplicitCast<Section>(insertAfterNode->GetAncestor(NodeType::Section));

        // Don't clone the content inside the section, we just want the properties of the section retained.
        auto cloneSection = System::ExplicitCast<Section>(currentSection->Clone(false));

        // However, make sure the clone section has a body but no empty first paragraph.
        cloneSection->EnsureMinimum();
        cloneSection->get_Body()->get_FirstParagraph()->Remove();

        insertAfterNode->get_Document()->InsertAfter(cloneSection, currentSection);

        // Append all nodes after the marker node to the new section. This will split the content at the section level at.
        // The marker so the sections from the other document can be inserted directly.
        SharedPtr<Node> currentNode = insertAfterNode->get_NextSibling();
        while (currentNode != nullptr)
        {
            SharedPtr<Node> nextNode = currentNode->get_NextSibling();
            cloneSection->get_Body()->AppendChild(currentNode);
            currentNode = nextNode;
        }

        // This object will be translating styles and lists during the import.
        auto importer = MakeObject<NodeImporter>(srcDoc, dstDoc, ImportFormatMode::UseDestinationStyles);

        for (const auto& srcSection : System::IterateOver<Section>(srcDoc->get_Sections()))
        {
            SharedPtr<Node> newNode = importer->ImportNode(srcSection, true);

            dstDoc->InsertAfter(newNode, currentSection);
            currentSection = System::ExplicitCast<Section>(newNode);
        }
    }
    //ExEnd:InsertDocumentWithSectionFormatting

    //ExStart:InsertDocumentAtMailMergeHandler
    //GistId:db2dfc4150d7c714bcac3782ae241d03
    class InsertDocumentAtMailMergeHandler : public IFieldMergingCallback
    {
    private:
        void FieldMerging(SharedPtr<FieldMergingArgs> args) override
        {
            if (args->get_DocumentFieldName() == u"Document_1")
            {
                // Use document builder to navigate to the merge field with the specified name.
                auto builder = MakeObject<DocumentBuilder>(args->get_Document());
                builder->MoveToMergeField(args->get_DocumentFieldName());

                // The name of the document to load and insert is stored in the field value.
                auto subDoc = MakeObject<Document>(System::ObjectExt::Unbox<String>(args->get_FieldValue()));

                InsertDocument(builder->get_CurrentParagraph(), subDoc);

                // The paragraph that contained the merge field might be empty now, and you probably want to delete it.
                if (!builder->get_CurrentParagraph()->get_HasChildNodes())
                {
                    builder->get_CurrentParagraph()->Remove();
                }

                // Indicate to the mail merge engine that we have inserted what we wanted.
                args->set_Text(nullptr);
            }
        }

        void ImageFieldMerging(SharedPtr<ImageFieldMergingArgs> args) override
        {
            ASPOSE_UNUSED(args);
            // Do nothing.
        }
    };
    //ExEnd:InsertDocumentAtMailMergeHandler

    //ExStart:InsertDocumentAtMailMergeBlobHandler
    class InsertDocumentAtMailMergeBlobHandler : public IFieldMergingCallback
    {
    private:
        /// <summary>
        /// This handler makes special processing for the "Document_1" field.
        /// The field value contains the path to load the document.
        /// We load the document and insert it into the current merge field.
        /// </summary>
        void FieldMerging(SharedPtr<FieldMergingArgs> e) override
        {
            if (e->get_DocumentFieldName() == u"Document_1")
            {
                auto builder = MakeObject<DocumentBuilder>(e->get_Document());
                builder->MoveToMergeField(e->get_DocumentFieldName());

                auto stream = MakeObject<System::IO::MemoryStream>(System::ExplicitCast<System::Array<uint8_t>>(e->get_FieldValue()));
                auto subDoc = MakeObject<Document>(stream);

                InsertDocument(builder->get_CurrentParagraph(), subDoc);

                // The paragraph that contained the merge field might be empty now, and you probably want to delete it.
                if (!builder->get_CurrentParagraph()->get_HasChildNodes())
                {
                    builder->get_CurrentParagraph()->Remove();
                }

                e->set_Text(nullptr);
            }
        }

        void ImageFieldMerging(SharedPtr<ImageFieldMergingArgs> args) override
        {
            ASPOSE_UNUSED(args);
            // Do nothing.
        }
    };
    //ExEnd:InsertDocumentAtMailMergeBlobHandler

    //ExStart:InsertDocumentAtReplaceHandler
    //GistId:db2dfc4150d7c714bcac3782ae241d03
    class InsertDocumentAtReplaceHandler : public IReplacingCallback
    {
    private:
        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            auto subDoc = MakeObject<Document>(MyDir + u"Document insertion 2.docx");

            // Insert a document after the paragraph, containing the match text.
            auto para = System::ExplicitCast<Paragraph>(args->get_MatchNode()->get_ParentNode());
            InsertDocument(para, subDoc);

            // Remove the paragraph with the match text.
            para->Remove();
            return ReplaceAction::Skip;
        }
    };
    //ExEnd:InsertDocumentAtReplaceHandler
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document
