#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Bookmark.h>
#include <Aspose.Words.Cpp/BookmarkCollection.h>
#include <Aspose.Words.Cpp/BookmarkEnd.h>
#include <Aspose.Words.Cpp/BookmarkStart.h>
#include <Aspose.Words.Cpp/CommentRangeEnd.h>
#include <Aspose.Words.Cpp/CommentRangeStart.h>
#include <Aspose.Words.Cpp/ControlChar.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Fields/FieldEnd.h>
#include <Aspose.Words.Cpp/Fields/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Fields/FieldSeparator.h>
#include <Aspose.Words.Cpp/Fields/FieldStart.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/VisitorAction.h>
#include <system/collections/list.h>
#include <system/exceptions.h>
#include <system/text/string_builder.h>

#include "DocsExamplesBase.h"
#include "Programming with Documents/Contents Management/Extract content helper.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Management {

class ExtractContent : public DocsExamplesBase
{
public:
    void ExtractContentBetweenBlockLevelNodes()
    {
        //ExStart:ExtractContentBetweenBlockLevelNodes
        auto doc = MakeObject<Document>(MyDir + u"Extract content.docx");

        auto startPara = System::ExplicitCast<Paragraph>(doc->get_LastSection()->GetChild(NodeType::Paragraph, 2, true));
        auto endTable = System::ExplicitCast<Table>(doc->get_LastSection()->GetChild(NodeType::Table, 0, true));

        // Extract the content between these nodes in the document. Include these markers in the extraction.
        SharedPtr<System::Collections::Generic::List<SharedPtr<Node>>> extractedNodes = ExtractContentHelper::ExtractContent(startPara, endTable, true);

        // Let's reverse the array to make inserting the content back into the document easier.
        extractedNodes->Reverse();

        while (extractedNodes->get_Count() > 0)
        {
            // Insert the last node from the reversed list.
            endTable->get_ParentNode()->InsertAfter(System::ExplicitCast<Node>(extractedNodes->idx_get(0)), endTable);
            // Remove this node from the list after insertion.
            extractedNodes->RemoveAt(0);
        }

        doc->Save(ArtifactsDir + u"ExtractContent.ExtractContentBetweenBlockLevelNodes.docx");
        //ExEnd:ExtractContentBetweenBlockLevelNodes
    }

    void ExtractContentBetweenBookmark()
    {
        //ExStart:ExtractContentBetweenBookmark
        auto doc = MakeObject<Document>(MyDir + u"Extract content.docx");

        SharedPtr<Section> section = doc->get_Sections()->idx_get(0);
        section->get_PageSetup()->set_LeftMargin(70.85);

        // Retrieve the bookmark from the document.
        SharedPtr<Bookmark> bookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"Bookmark1");
        // We use the BookmarkStart and BookmarkEnd nodes as markers.
        SharedPtr<BookmarkStart> bookmarkStart = bookmark->get_BookmarkStart();
        SharedPtr<BookmarkEnd> bookmarkEnd = bookmark->get_BookmarkEnd();

        // Firstly, extract the content between these nodes, including the bookmark.
        SharedPtr<System::Collections::Generic::List<SharedPtr<Node>>> extractedNodesInclusive =
            ExtractContentHelper::ExtractContent(bookmarkStart, bookmarkEnd, true);

        SharedPtr<Document> dstDoc = ExtractContentHelper::GenerateDocument(doc, extractedNodesInclusive);
        dstDoc->Save(ArtifactsDir + u"ExtractContent.ExtractContentBetweenBookmark.IncludingBookmark.docx");

        // Secondly, extract the content between these nodes this time without including the bookmark.
        SharedPtr<System::Collections::Generic::List<SharedPtr<Node>>> extractedNodesExclusive =
            ExtractContentHelper::ExtractContent(bookmarkStart, bookmarkEnd, false);

        dstDoc = ExtractContentHelper::GenerateDocument(doc, extractedNodesExclusive);
        dstDoc->Save(ArtifactsDir + u"ExtractContent.ExtractContentBetweenBookmark.WithoutBookmark.docx");
        //ExEnd:ExtractContentBetweenBookmark
    }

    void ExtractContentBetweenCommentRange()
    {
        //ExStart:ExtractContentBetweenCommentRange
        auto doc = MakeObject<Document>(MyDir + u"Extract content.docx");

        // This is a quick way of getting both comment nodes.
        // Your code should have a proper method of retrieving each corresponding start and end node.
        auto commentStart = System::ExplicitCast<CommentRangeStart>(doc->GetChild(NodeType::CommentRangeStart, 0, true));
        auto commentEnd = System::ExplicitCast<CommentRangeEnd>(doc->GetChild(NodeType::CommentRangeEnd, 0, true));

        // Firstly, extract the content between these nodes including the comment as well.
        SharedPtr<System::Collections::Generic::List<SharedPtr<Node>>> extractedNodesInclusive =
            ExtractContentHelper::ExtractContent(commentStart, commentEnd, true);

        SharedPtr<Document> dstDoc = ExtractContentHelper::GenerateDocument(doc, extractedNodesInclusive);
        dstDoc->Save(ArtifactsDir + u"ExtractContent.ExtractContentBetweenCommentRange.IncludingComment.docx");

        // Secondly, extract the content between these nodes without the comment.
        SharedPtr<System::Collections::Generic::List<SharedPtr<Node>>> extractedNodesExclusive =
            ExtractContentHelper::ExtractContent(commentStart, commentEnd, false);

        dstDoc = ExtractContentHelper::GenerateDocument(doc, extractedNodesExclusive);
        dstDoc->Save(ArtifactsDir + u"ExtractContent.ExtractContentBetweenCommentRange.WithoutComment.docx");
        //ExEnd:ExtractContentBetweenCommentRange
    }

    void ExtractContentBetweenParagraphs()
    {
        //ExStart:ExtractContentBetweenParagraphs
        auto doc = MakeObject<Document>(MyDir + u"Extract content.docx");

        auto startPara = System::ExplicitCast<Paragraph>(doc->get_FirstSection()->get_Body()->GetChild(NodeType::Paragraph, 6, true));
        auto endPara = System::ExplicitCast<Paragraph>(doc->get_FirstSection()->get_Body()->GetChild(NodeType::Paragraph, 10, true));

        // Extract the content between these nodes in the document. Include these markers in the extraction.
        SharedPtr<System::Collections::Generic::List<SharedPtr<Node>>> extractedNodes = ExtractContentHelper::ExtractContent(startPara, endPara, true);

        SharedPtr<Document> dstDoc = ExtractContentHelper::GenerateDocument(doc, extractedNodes);
        dstDoc->Save(ArtifactsDir + u"ExtractContent.ExtractContentBetweenParagraphs.docx");
        //ExEnd:ExtractContentBetweenParagraphs
    }

    void ExtractContentBetweenParagraphStyles()
    {
        //ExStart:ExtractContentBetweenParagraphStyles
        auto doc = MakeObject<Document>(MyDir + u"Extract content.docx");

        // Gather a list of the paragraphs using the respective heading styles.
        SharedPtr<System::Collections::Generic::List<SharedPtr<Paragraph>>> parasStyleHeading1 = ExtractContentHelper::ParagraphsByStyleName(doc, u"Heading 1");
        SharedPtr<System::Collections::Generic::List<SharedPtr<Paragraph>>> parasStyleHeading3 = ExtractContentHelper::ParagraphsByStyleName(doc, u"Heading 3");

        // Use the first instance of the paragraphs with those styles.
        SharedPtr<Node> startPara1 = parasStyleHeading1->idx_get(0);
        SharedPtr<Node> endPara1 = parasStyleHeading3->idx_get(0);

        // Extract the content between these nodes in the document. Don't include these markers in the extraction.
        SharedPtr<System::Collections::Generic::List<SharedPtr<Node>>> extractedNodes = ExtractContentHelper::ExtractContent(startPara1, endPara1, false);

        SharedPtr<Document> dstDoc = ExtractContentHelper::GenerateDocument(doc, extractedNodes);
        dstDoc->Save(ArtifactsDir + u"ExtractContent.ExtractContentBetweenParagraphStyles.docx");
        //ExEnd:ExtractContentBetweenParagraphStyles
    }

    void ExtractContentBetweenRuns()
    {
        //ExStart:ExtractContentBetweenRuns
        auto doc = MakeObject<Document>(MyDir + u"Extract content.docx");

        auto para = System::ExplicitCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 7, true));

        SharedPtr<Run> startRun = para->get_Runs()->idx_get(1);
        SharedPtr<Run> endRun = para->get_Runs()->idx_get(4);

        // Extract the content between these nodes in the document. Include these markers in the extraction.
        SharedPtr<System::Collections::Generic::List<SharedPtr<Node>>> extractedNodes = ExtractContentHelper::ExtractContent(startRun, endRun, true);

        auto node = System::ExplicitCast<Node>(extractedNodes->idx_get(0));
        std::cout << node->ToString(SaveFormat::Text) << std::endl;
        //ExEnd:ExtractContentBetweenRuns
    }

    void ExtractContentUsingDocumentVisitor()
    {
        //ExStart:ExtractContentUsingDocumentVisitor
        auto doc = MakeObject<Document>(MyDir + u"Absolute position tab.docx");

        auto myConverter = MakeObject<ExtractContent::MyDocToTxtWriter>();

        // This is the well known Visitor pattern. Get the model to accept a visitor.
        // The model will iterate through itself by calling the corresponding methods.
        // On the visitor object (this is called visiting).
        // Note that every node in the object model has the accept method so the visiting
        // can be executed not only for the whole document, but for any node in the document.
        doc->Accept(myConverter);

        // Once the visiting is complete, we can retrieve the result of the operation,
        // That in this example, has accumulated in the visitor.
        std::cout << myConverter->GetText() << std::endl;
        //ExEnd:ExtractContentUsingDocumentVisitor
    }

    //ExStart:MyDocToTxtWriter
    /// <summary>
    /// Simple implementation of saving a document in the plain text format. Implemented as a Visitor.
    /// </summary>
    class MyDocToTxtWriter : public DocumentVisitor
    {
    public:
        MyDocToTxtWriter() : mIsSkipText(false)
        {
            mIsSkipText = false;
            mBuilder = MakeObject<System::Text::StringBuilder>();
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        String GetText()
        {
            return mBuilder->ToString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        VisitorAction VisitRun(SharedPtr<Run> run) override
        {
            AppendText(run->get_Text());

            // Let the visitor continue visiting other nodes.
            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        VisitorAction VisitFieldStart(SharedPtr<FieldStart> fieldStart) override
        {
            ASPOSE_UNUSED(fieldStart);
            // In Microsoft Word, a field code (such as "MERGEFIELD FieldName") follows
            // after a field start character. We want to skip field codes and output field.
            // Result only, therefore we use a flag to suspend the output while inside a field code.
            // Note this is a very simplistic implementation and will not work very well.
            // If you have nested fields in a document.
            mIsSkipText = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        VisitorAction VisitFieldSeparator(SharedPtr<FieldSeparator> fieldSeparator) override
        {
            ASPOSE_UNUSED(fieldSeparator);
            // Once reached a field separator node, we enable the output because we are
            // now entering the field result nodes.
            mIsSkipText = false;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        VisitorAction VisitFieldEnd(SharedPtr<FieldEnd> fieldEnd) override
        {
            ASPOSE_UNUSED(fieldEnd);
            // Make sure we enable the output when reached a field end because some fields
            // do not have field separator and do not have field result.
            mIsSkipText = false;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when visiting of a Paragraph node is ended in the document.
        /// </summary>
        VisitorAction VisitParagraphEnd(SharedPtr<Paragraph> paragraph) override
        {
            ASPOSE_UNUSED(paragraph);
            // When outputting to plain text we output Cr+Lf characters.
            AppendText(ControlChar::CrLf());

            return VisitorAction::Continue;
        }

        VisitorAction VisitBodyStart(SharedPtr<Body> body) override
        {
            ASPOSE_UNUSED(body);
            // We can detect beginning and end of all composite nodes such as Section, Body,
            // Table, Paragraph etc and provide custom handling for them.
            mBuilder->Append(u"*** Body Started ***\r\n");

            return VisitorAction::Continue;
        }

        VisitorAction VisitBodyEnd(SharedPtr<Body> body) override
        {
            ASPOSE_UNUSED(body);
            mBuilder->Append(u"*** Body Ended ***\r\n");
            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a HeaderFooter node is encountered in the document.
        /// </summary>
        VisitorAction VisitHeaderFooterStart(SharedPtr<HeaderFooter> headerFooter) override
        {
            ASPOSE_UNUSED(headerFooter);
            // Returning this value from a visitor method causes visiting of this
            // Node to stop and move on to visiting the next sibling node
            // The net effect in this example is that the text of headers and footers
            // Is not included in the resulting output
            return VisitorAction::SkipThisNode;
        }

    private:
        SharedPtr<System::Text::StringBuilder> mBuilder;
        bool mIsSkipText;

        /// <summary>
        /// Adds text to the current output. Honors the enabled/disabled output flag.
        /// </summary>
        void AppendText(String text)
        {
            if (!mIsSkipText)
            {
                mBuilder->Append(text);
            }
        }
    };
    //ExEnd:MyDocToTxtWriter

    void ExtractContentUsingField()
    {
        //ExStart:ExtractContentUsingField
        auto doc = MakeObject<Document>(MyDir + u"Extract content.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
        // We could also get FieldStarts of a field using GetChildNode method as in the other examples.
        builder->MoveToMergeField(u"Fullname", false, false);

        // The builder cursor should be positioned at the start of the field.
        auto startField = System::ExplicitCast<FieldStart>(builder->get_CurrentNode());
        auto endPara = System::ExplicitCast<Paragraph>(doc->get_FirstSection()->GetChild(NodeType::Paragraph, 5, true));

        // Extract the content between these nodes in the document. Don't include these markers in the extraction.
        SharedPtr<System::Collections::Generic::List<SharedPtr<Node>>> extractedNodes = ExtractContentHelper::ExtractContent(startField, endPara, false);

        SharedPtr<Document> dstDoc = ExtractContentHelper::GenerateDocument(doc, extractedNodes);
        dstDoc->Save(ArtifactsDir + u"ExtractContent.ExtractContentUsingField.docx");
        //ExEnd:ExtractContentUsingField
    }

    void ExtractTableOfContents()
    {
        auto doc = MakeObject<Document>(MyDir + u"Table of contents.docx");

        for (const auto& field : System::IterateOver(doc->get_Range()->get_Fields()))
        {
            if (field->get_Type() == FieldType::FieldHyperlink)
            {
                auto hyperlink = System::ExplicitCast<FieldHyperlink>(field);
                if (hyperlink->get_SubAddress() != nullptr && hyperlink->get_SubAddress().StartsWith(u"_Toc"))
                {
                    auto tocItem = System::ExplicitCast<Paragraph>(field->get_Start()->GetAncestor(NodeType::Paragraph));

                    std::cout << tocItem->ToString(SaveFormat::Text).Trim() << std::endl;
                    std::cout << "------------------" << std::endl;

                    SharedPtr<Bookmark> bm = doc->get_Range()->get_Bookmarks()->idx_get(hyperlink->get_SubAddress());
                    auto pointer = System::ExplicitCast<Paragraph>(bm->get_BookmarkStart()->GetAncestor(NodeType::Paragraph));

                    std::cout << pointer->ToString(SaveFormat::Text) << std::endl;
                }
            }
        }
    }

    void ExtractTextOnly()
    {
        //ExStart:ExtractTextOnly
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u"MERGEFIELD Field");

        std::cout << (String(u"GetText() Result: ") + doc->GetText()) << std::endl;

        // When converted to text it will not retrieve fields code or special characters,
        // but will still contain some natural formatting characters such as paragraph markers etc.
        // This is the same as "viewing" the document as if it was opened in a text editor.
        std::cout << (String(u"ToString() Result: ") + doc->ToString(SaveFormat::Text)) << std::endl;
        //ExEnd:ExtractTextOnly
    }

    void ExtractContentBasedOnStyles()
    {
        //ExStart:ExtractContentBasedOnStyles
        auto doc = MakeObject<Document>(MyDir + u"Styles.docx");

        const String paraStyle = u"Heading 1";
        const String runStyle = u"Intense Emphasis";

        SharedPtr<System::Collections::Generic::List<SharedPtr<Paragraph>>> paragraphs = ParagraphsByStyleName(doc, paraStyle);
        std::cout << "Paragraphs with \"" << paraStyle << "\" styles (" << paragraphs->get_Count() << "):" << std::endl;

        for (const auto& paragraph : paragraphs)
        {
            std::cout << paragraph->ToString(SaveFormat::Text);
        }

        SharedPtr<System::Collections::Generic::List<SharedPtr<Run>>> runs = RunsByStyleName(doc, runStyle);
        std::cout << "\nRuns with \"" << runStyle << "\" styles (" << runs->get_Count() << "):" << std::endl;

        for (const auto& run : runs)
        {
            std::cout << run->get_Range()->get_Text() << std::endl;
        }
        //ExEnd:ExtractContentBasedOnStyles
    }

    //ExStart:ParagraphsByStyleName
    SharedPtr<System::Collections::Generic::List<SharedPtr<Paragraph>>> ParagraphsByStyleName(SharedPtr<Document> doc, String styleName)
    {
        SharedPtr<System::Collections::Generic::List<SharedPtr<Paragraph>>> paragraphsWithStyle =
            MakeObject<System::Collections::Generic::List<SharedPtr<Paragraph>>>();
        SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(NodeType::Paragraph, true);

        for (const auto& paragraph : System::IterateOver<Paragraph>(paragraphs))
        {
            if (paragraph->get_ParagraphFormat()->get_Style()->get_Name() == styleName)
            {
                paragraphsWithStyle->Add(paragraph);
            }
        }

        return paragraphsWithStyle;
    }
    //ExEnd:ParagraphsByStyleName

    //ExStart:RunsByStyleName
    SharedPtr<System::Collections::Generic::List<SharedPtr<Run>>> RunsByStyleName(SharedPtr<Document> doc, String styleName)
    {
        SharedPtr<System::Collections::Generic::List<SharedPtr<Run>>> runsWithStyle = MakeObject<System::Collections::Generic::List<SharedPtr<Run>>>();
        SharedPtr<NodeCollection> runs = doc->GetChildNodes(NodeType::Run, true);

        for (const auto& run : System::IterateOver<Aspose::Words::Run>(runs))
        {
            if (run->get_Font()->get_Style()->get_Name() == styleName)
            {
                runsWithStyle->Add(run);
            }
        }

        return runsWithStyle;
    }
    //ExEnd:RunsByStyleName

    void ExtractPrintText()
    {
        //ExStart:ExtractText
        //GistId:458eb4fd5bd1de8b06fab4d1ef1acdc6
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::ExplicitCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        // The range text will include control characters such as "\a" for a cell.
        // You can call ToString and pass SaveFormat.Text on the desired node to find the plain text content.

        std::cout << "Contents of the table: " << std::endl;
        std::cout << table->get_Range()->get_Text() << std::endl;
        //ExEnd:ExtractText

        //ExStart:PrintTextRangeRowAndTable
        //GistId:458eb4fd5bd1de8b06fab4d1ef1acdc6
        std::cout << "\nContents of the row: " << std::endl;
        std::cout << table->get_Rows()->idx_get(1)->get_Range()->get_Text() << std::endl;

        std::cout << "\nContents of the cell: " << std::endl;
        std::cout << table->get_LastRow()->get_LastCell()->get_Range()->get_Text() << std::endl;
        //ExEnd:PrintTextRangeRowAndTable
    }

    void ExtractImagesToFiles()
    {
        //ExStart:ExtractImagesToFiles
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        SharedPtr<NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);
        int imageIndex = 0;

        for (const auto& shape : System::IterateOver<Shape>(shapes))
        {
            if (shape->get_HasImage())
            {
                String imageFileName =
                    String::Format(u"Image.ExportImages.{0}_{1}", imageIndex, FileFormatUtil::ImageTypeToExtension(shape->get_ImageData()->get_ImageType()));

                shape->get_ImageData()->Save(ArtifactsDir + imageFileName);
                imageIndex++;
            }
        }
        //ExEnd:ExtractImagesToFiles
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Contents_Management
