#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/SubDocument.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRange.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeEnd.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeStart.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Markup/SmartTag.h>
#include <Aspose.Words.Cpp/Model/Math/MathObjectType.h>
#include <Aspose.Words.Cpp/Model/Math/OfficeMath.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeStart.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <system/array.h>
#include <system/date_time.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/string_builder.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Math;
using namespace Aspose::Words::Tables;

namespace ApiExamples {

class ExDocumentVisitor : public ApiExampleBase
{
public:
    //ExStart
    //ExFor:Document.Accept(DocumentVisitor)
    //ExFor:Body.Accept(DocumentVisitor)
    //ExFor:SubDocument.Accept(DocumentVisitor)
    //ExFor:DocumentVisitor
    //ExFor:DocumentVisitor.VisitRun(Run)
    //ExFor:DocumentVisitor.VisitDocumentEnd(Document)
    //ExFor:DocumentVisitor.VisitDocumentStart(Document)
    //ExFor:DocumentVisitor.VisitSectionEnd(Section)
    //ExFor:DocumentVisitor.VisitSectionStart(Section)
    //ExFor:DocumentVisitor.VisitBodyStart(Body)
    //ExFor:DocumentVisitor.VisitBodyEnd(Body)
    //ExFor:DocumentVisitor.VisitParagraphStart(Paragraph)
    //ExFor:DocumentVisitor.VisitParagraphEnd(Paragraph)
    //ExFor:DocumentVisitor.VisitSubDocument(SubDocument)
    //ExSummary:Shows how to use a document visitor to print a document's node structure.
    void DocStructureToText()
    {
        auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");
        auto visitor = MakeObject<ExDocumentVisitor::DocStructurePrinter>();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all of the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc->Accept(visitor);

        std::cout << visitor->GetText() << std::endl;
        TestDocStructureToText(visitor);
        //ExSkip
    }

    /// <summary>
    /// Traverses a node's tree of child nodes, and creates a map of this tree in the form of a string.
    /// </summary>
    class DocStructurePrinter : public DocumentVisitor
    {
    public:
        DocStructurePrinter() : mDocTraversalDepth(0)
        {
            mAcceptingNodeChildTree = MakeObject<System::Text::StringBuilder>();
        }

        String GetText()
        {
            return mAcceptingNodeChildTree->ToString();
        }

        /// <summary>
        /// Called when a Document node is encountered.
        /// </summary>
        VisitorAction VisitDocumentStart(SharedPtr<Document> doc) override
        {
            int childNodeCount = doc->GetChildNodes(NodeType::Any, true)->get_Count();

            IndentAndAppendLine(String(u"[Document start] Child nodes: ") + childNodeCount);
            mDocTraversalDepth++;

            // Allow the visitor to continue visiting other nodes.
            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of a Document node have been visited.
        /// </summary>
        VisitorAction VisitDocumentEnd(SharedPtr<Document> doc) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[Document end]");

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Section node is encountered in the document.
        /// </summary>
        VisitorAction VisitSectionStart(SharedPtr<Section> section) override
        {
            // Get the index of our section within the document
            SharedPtr<NodeCollection> docSections = section->get_Document()->GetChildNodes(NodeType::Section, false);
            int sectionIndex = docSections->IndexOf(section);

            IndentAndAppendLine(String(u"[Section start] Section index: ") + sectionIndex);
            mDocTraversalDepth++;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of a Section node have been visited.
        /// </summary>
        VisitorAction VisitSectionEnd(SharedPtr<Section> section) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[Section end]");

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Body node is encountered in the document.
        /// </summary>
        VisitorAction VisitBodyStart(SharedPtr<Body> body) override
        {
            int paragraphCount = body->get_Paragraphs()->get_Count();
            IndentAndAppendLine(String(u"[Body start] Paragraphs: ") + paragraphCount);
            mDocTraversalDepth++;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of a Body node have been visited.
        /// </summary>
        VisitorAction VisitBodyEnd(SharedPtr<Body> body) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[Body end]");

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Paragraph node is encountered in the document.
        /// </summary>
        VisitorAction VisitParagraphStart(SharedPtr<Paragraph> paragraph) override
        {
            IndentAndAppendLine(u"[Paragraph start]");
            mDocTraversalDepth++;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of a Paragraph node have been visited.
        /// </summary>
        VisitorAction VisitParagraphEnd(SharedPtr<Paragraph> paragraph) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[Paragraph end]");

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        VisitorAction VisitRun(SharedPtr<Run> run) override
        {
            IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a SubDocument node is encountered in the document.
        /// </summary>
        VisitorAction VisitSubDocument(SharedPtr<SubDocument> subDocument) override
        {
            IndentAndAppendLine(u"[SubDocument]");

            return VisitorAction::Continue;
        }

    private:
        int mDocTraversalDepth;
        SharedPtr<System::Text::StringBuilder> mAcceptingNodeChildTree;

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                mAcceptingNodeChildTree->Append(u"|  ");
            }

            mAcceptingNodeChildTree->AppendLine(text);
        }
    };
    //ExEnd

protected:
    void TestDocStructureToText(SharedPtr<ExDocumentVisitor::DocStructurePrinter> visitor)
    {
        String visitorText = visitor->GetText();

        ASSERT_TRUE(visitorText.Contains(u"[Document start]"));
        ASSERT_TRUE(visitorText.Contains(u"[Document end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Section start]"));
        ASSERT_TRUE(visitorText.Contains(u"[Section end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Body start]"));
        ASSERT_TRUE(visitorText.Contains(u"[Body end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Paragraph start]"));
        ASSERT_TRUE(visitorText.Contains(u"[Paragraph end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Run]"));
        ASSERT_TRUE(visitorText.Contains(u"[SubDocument]"));
    }

public:
    //ExStart
    //ExFor:Cell.Accept(DocumentVisitor)
    //ExFor:Cell.IsFirstCell
    //ExFor:Cell.IsLastCell
    //ExFor:DocumentVisitor.VisitTableEnd(Tables.Table)
    //ExFor:DocumentVisitor.VisitTableStart(Tables.Table)
    //ExFor:DocumentVisitor.VisitRowEnd(Tables.Row)
    //ExFor:DocumentVisitor.VisitRowStart(Tables.Row)
    //ExFor:DocumentVisitor.VisitCellStart(Tables.Cell)
    //ExFor:DocumentVisitor.VisitCellEnd(Tables.Cell)
    //ExFor:Row.Accept(DocumentVisitor)
    //ExFor:Row.FirstCell
    //ExFor:Row.GetText
    //ExFor:Row.IsFirstRow
    //ExFor:Row.LastCell
    //ExFor:Row.ParentTable
    //ExSummary:Shows how to print the node structure of every table in a document.
    void TableToText()
    {
        auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");
        auto visitor = MakeObject<ExDocumentVisitor::TableStructurePrinter>();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all of the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc->Accept(visitor);

        std::cout << visitor->GetText() << std::endl;
        TestTableToText(visitor);
        //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered Table nodes and their children.
    /// </summary>
    class TableStructurePrinter : public DocumentVisitor
    {
    public:
        TableStructurePrinter() : mVisitorIsInsideTable(false), mDocTraversalDepth(0)
        {
            mVisitedTables = MakeObject<System::Text::StringBuilder>();
            mVisitorIsInsideTable = false;
        }

        String GetText()
        {
            return mVisitedTables->ToString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// Runs that are not within tables are not recorded.
        /// </summary>
        VisitorAction VisitRun(SharedPtr<Run> run) override
        {
            if (mVisitorIsInsideTable)
            {
                IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Table is encountered in the document.
        /// </summary>
        VisitorAction VisitTableStart(SharedPtr<Table> table) override
        {
            int rows = 0;
            int columns = 0;

            if (table->get_Rows()->get_Count() > 0)
            {
                rows = table->get_Rows()->get_Count();
                columns = table->get_FirstRow()->get_Count();
            }

            IndentAndAppendLine(String(u"[Table start] Size: ") + rows + u"x" + columns);
            mDocTraversalDepth++;
            mVisitorIsInsideTable = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of a Table node have been visited.
        /// </summary>
        VisitorAction VisitTableEnd(SharedPtr<Table> table) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[Table end]");
            mVisitorIsInsideTable = false;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Row node is encountered in the document.
        /// </summary>
        VisitorAction VisitRowStart(SharedPtr<Row> row) override
        {
            String rowContents = row->GetText().TrimEnd(MakeArray<char16_t>({u'\x0007', u' '})).Replace(u"\u0007", u", ");
            int rowWidth = row->IndexOf(row->get_LastCell()) + 1;
            int rowIndex = row->get_ParentTable()->IndexOf(row);
            String rowStatusInTable = row->get_IsFirstRow() && row->get_IsLastRow() ? u"only"
                                              : row->get_IsFirstRow()                       ? u"first"
                                              : row->get_IsLastRow()                        ? String(u"last")
                                                                                            : String(u"");
            if (rowStatusInTable != u"")
            {
                rowStatusInTable = String::Format(u", the {0} row in this table,", rowStatusInTable);
            }

            IndentAndAppendLine(String::Format(u"[Row start] Row #{0}{1} width {2}, \"{3}\"", ++rowIndex, rowStatusInTable, rowWidth, rowContents));
            mDocTraversalDepth++;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of a Row node have been visited.
        /// </summary>
        VisitorAction VisitRowEnd(SharedPtr<Row> row) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[Row end]");

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Cell node is encountered in the document.
        /// </summary>
        VisitorAction VisitCellStart(SharedPtr<Cell> cell) override
        {
            SharedPtr<Row> row = cell->get_ParentRow();
            SharedPtr<Table> table = row->get_ParentTable();
            String cellStatusInRow = cell->get_IsFirstCell() && cell->get_IsLastCell() ? u"only"
                                             : cell->get_IsFirstCell()                         ? u"first"
                                             : cell->get_IsLastCell()                          ? String(u"last")
                                                                                               : String(u"");
            if (cellStatusInRow != u"")
            {
                cellStatusInRow = String::Format(u", the {0} cell in this row", cellStatusInRow);
            }

            IndentAndAppendLine(String::Format(u"[Cell start] Row {0}, Col {1}{2}", table->IndexOf(row) + 1, row->IndexOf(cell) + 1, cellStatusInRow));
            mDocTraversalDepth++;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of a Cell node have been visited.
        /// </summary>
        VisitorAction VisitCellEnd(SharedPtr<Cell> cell) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[Cell end]");
            return VisitorAction::Continue;
        }

    private:
        bool mVisitorIsInsideTable;
        int mDocTraversalDepth;
        SharedPtr<System::Text::StringBuilder> mVisitedTables;

        /// <summary>
        /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is
        /// into the current table's tree of child nodes.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                mVisitedTables->Append(u"|  ");
            }

            mVisitedTables->AppendLine(text);
        }
    };
    //ExEnd

protected:
    void TestTableToText(SharedPtr<ExDocumentVisitor::TableStructurePrinter> visitor)
    {
        String visitorText = visitor->GetText();

        ASSERT_TRUE(visitorText.Contains(u"[Table start]"));
        ASSERT_TRUE(visitorText.Contains(u"[Table end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Row start]"));
        ASSERT_TRUE(visitorText.Contains(u"[Row end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Cell start]"));
        ASSERT_TRUE(visitorText.Contains(u"[Cell end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Run]"));
    }

public:
    //ExStart
    //ExFor:DocumentVisitor.VisitCommentStart(Comment)
    //ExFor:DocumentVisitor.VisitCommentEnd(Comment)
    //ExFor:DocumentVisitor.VisitCommentRangeEnd(CommentRangeEnd)
    //ExFor:DocumentVisitor.VisitCommentRangeStart(CommentRangeStart)
    //ExSummary:Shows how to print the node structure of every comment and comment range in a document.
    void CommentsToText()
    {
        auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");
        auto visitor = MakeObject<ExDocumentVisitor::CommentStructurePrinter>();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all of the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc->Accept(visitor);

        std::cout << visitor->GetText() << std::endl;
        TestCommentsToText(visitor);
        //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered Comment/CommentRange nodes and their children.
    /// </summary>
    class CommentStructurePrinter : public DocumentVisitor
    {
    public:
        CommentStructurePrinter() : mVisitorIsInsideComment(false), mDocTraversalDepth(0)
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
            mVisitorIsInsideComment = false;
        }

        String GetText()
        {
            return mBuilder->ToString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// A Run is only recorded if it is a child of a Comment or CommentRange node.
        /// </summary>
        VisitorAction VisitRun(SharedPtr<Run> run) override
        {
            if (mVisitorIsInsideComment)
            {
                IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a CommentRangeStart node is encountered in the document.
        /// </summary>
        VisitorAction VisitCommentRangeStart(SharedPtr<CommentRangeStart> commentRangeStart) override
        {
            IndentAndAppendLine(String(u"[Comment range start] ID: ") + commentRangeStart->get_Id());
            mDocTraversalDepth++;
            mVisitorIsInsideComment = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a CommentRangeEnd node is encountered in the document.
        /// </summary>
        VisitorAction VisitCommentRangeEnd(SharedPtr<CommentRangeEnd> commentRangeEnd) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[Comment range end]");
            mVisitorIsInsideComment = false;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Comment node is encountered in the document.
        /// </summary>
        VisitorAction VisitCommentStart(SharedPtr<Comment> comment) override
        {
            IndentAndAppendLine(String::Format(u"[Comment start] For comment range ID {0}, By {1} on {2}", comment->get_Id(), comment->get_Author(),
                                                       comment->get_DateTime()));
            mDocTraversalDepth++;
            mVisitorIsInsideComment = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of a Comment node have been visited.
        /// </summary>
        VisitorAction VisitCommentEnd(SharedPtr<Comment> comment) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[Comment end]");
            mVisitorIsInsideComment = false;

            return VisitorAction::Continue;
        }

    private:
        bool mVisitorIsInsideComment;
        int mDocTraversalDepth;
        SharedPtr<System::Text::StringBuilder> mBuilder;

        /// <summary>
        /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is
        /// into a comment/comment range's tree of child nodes.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                mBuilder->Append(u"|  ");
            }

            mBuilder->AppendLine(text);
        }
    };
    //ExEnd

protected:
    void TestCommentsToText(SharedPtr<ExDocumentVisitor::CommentStructurePrinter> visitor)
    {
        String visitorText = visitor->GetText();

        ASSERT_TRUE(visitorText.Contains(u"[Comment range start]"));
        ASSERT_TRUE(visitorText.Contains(u"[Comment range end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Comment start]"));
        ASSERT_TRUE(visitorText.Contains(u"[Comment end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Run]"));
    }

public:
    //ExStart
    //ExFor:DocumentVisitor.VisitFieldStart
    //ExFor:DocumentVisitor.VisitFieldEnd
    //ExFor:DocumentVisitor.VisitFieldSeparator
    //ExSummary:Shows how to print the node structure of every field in a document.
    void FieldToText()
    {
        auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");
        auto visitor = MakeObject<ExDocumentVisitor::FieldStructurePrinter>();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all of the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc->Accept(visitor);

        std::cout << visitor->GetText() << std::endl;
        TestFieldToText(visitor);
        //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered Field nodes and their children.
    /// </summary>
    class FieldStructurePrinter : public DocumentVisitor
    {
    public:
        FieldStructurePrinter() : mVisitorIsInsideField(false), mDocTraversalDepth(0)
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
            mVisitorIsInsideField = false;
        }

        String GetText()
        {
            return mBuilder->ToString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        VisitorAction VisitRun(SharedPtr<Run> run) override
        {
            if (mVisitorIsInsideField)
            {
                IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        VisitorAction VisitFieldStart(SharedPtr<FieldStart> fieldStart) override
        {
            IndentAndAppendLine(String(u"[Field start] FieldType: ") + System::ObjectExt::ToString(fieldStart->get_FieldType()));
            mDocTraversalDepth++;
            mVisitorIsInsideField = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        VisitorAction VisitFieldEnd(SharedPtr<FieldEnd> fieldEnd) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[Field end]");
            mVisitorIsInsideField = false;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        VisitorAction VisitFieldSeparator(SharedPtr<FieldSeparator> fieldSeparator) override
        {
            IndentAndAppendLine(u"[FieldSeparator]");

            return VisitorAction::Continue;
        }

    private:
        bool mVisitorIsInsideField;
        int mDocTraversalDepth;
        SharedPtr<System::Text::StringBuilder> mBuilder;

        /// <summary>
        /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is
        /// into the field's tree of child nodes.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                mBuilder->Append(u"|  ");
            }

            mBuilder->AppendLine(text);
        }
    };
    //ExEnd

protected:
    void TestFieldToText(SharedPtr<ExDocumentVisitor::FieldStructurePrinter> visitor)
    {
        String visitorText = visitor->GetText();

        ASSERT_TRUE(visitorText.Contains(u"[Field start]"));
        ASSERT_TRUE(visitorText.Contains(u"[Field end]"));
        ASSERT_TRUE(visitorText.Contains(u"[FieldSeparator]"));
        ASSERT_TRUE(visitorText.Contains(u"[Run]"));
    }

public:
    //ExStart
    //ExFor:DocumentVisitor.VisitHeaderFooterStart(HeaderFooter)
    //ExFor:DocumentVisitor.VisitHeaderFooterEnd(HeaderFooter)
    //ExFor:HeaderFooter.Accept(DocumentVisitor)
    //ExFor:HeaderFooterCollection.ToArray
    //ExFor:Run.Accept(DocumentVisitor)
    //ExFor:Run.GetText
    //ExSummary:Shows how to print the node structure of every header and footer in a document.
    void HeaderFooterToText()
    {
        auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");
        auto visitor = MakeObject<ExDocumentVisitor::HeaderFooterStructurePrinter>();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all of the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc->Accept(visitor);

        std::cout << visitor->GetText() << std::endl;

        // An alternative way of accessing a document's header/footers section-by-section is by accessing the collection.
        ArrayPtr<SharedPtr<HeaderFooter>> headerFooters = doc->get_FirstSection()->get_HeadersFooters()->ToArray();
        ASSERT_EQ(3, headerFooters->get_Length());
        TestHeaderFooterToText(visitor);
        //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered HeaderFooter nodes and their children.
    /// </summary>
    class HeaderFooterStructurePrinter : public DocumentVisitor
    {
    public:
        HeaderFooterStructurePrinter() : mVisitorIsInsideHeaderFooter(false), mDocTraversalDepth(0)
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
            mVisitorIsInsideHeaderFooter = false;
        }

        String GetText()
        {
            return mBuilder->ToString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        VisitorAction VisitRun(SharedPtr<Run> run) override
        {
            if (mVisitorIsInsideHeaderFooter)
            {
                IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a HeaderFooter node is encountered in the document.
        /// </summary>
        VisitorAction VisitHeaderFooterStart(SharedPtr<HeaderFooter> headerFooter) override
        {
            IndentAndAppendLine(String(u"[HeaderFooter start] HeaderFooterType: ") + System::ObjectExt::ToString(headerFooter->get_HeaderFooterType()));
            mDocTraversalDepth++;
            mVisitorIsInsideHeaderFooter = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of a HeaderFooter node have been visited.
        /// </summary>
        VisitorAction VisitHeaderFooterEnd(SharedPtr<HeaderFooter> headerFooter) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[HeaderFooter end]");
            mVisitorIsInsideHeaderFooter = false;

            return VisitorAction::Continue;
        }

    private:
        bool mVisitorIsInsideHeaderFooter;
        int mDocTraversalDepth;
        SharedPtr<System::Text::StringBuilder> mBuilder;

        /// <summary>
        /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                mBuilder->Append(u"|  ");
            }

            mBuilder->AppendLine(text);
        }
    };
    //ExEnd

protected:
    void TestHeaderFooterToText(SharedPtr<ExDocumentVisitor::HeaderFooterStructurePrinter> visitor)
    {
        String visitorText = visitor->GetText();

        ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter start] HeaderFooterType: HeaderPrimary"));
        ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter end]"));
        ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter start] HeaderFooterType: HeaderFirst"));
        ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter start] HeaderFooterType: HeaderEven"));
        ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter start] HeaderFooterType: FooterPrimary"));
        ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter start] HeaderFooterType: FooterFirst"));
        ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter start] HeaderFooterType: FooterEven"));
        ASSERT_TRUE(visitorText.Contains(u"[Run]"));
    }

public:
    //ExStart
    //ExFor:DocumentVisitor.VisitEditableRangeEnd(EditableRangeEnd)
    //ExFor:DocumentVisitor.VisitEditableRangeStart(EditableRangeStart)
    //ExSummary:Shows how to print the node structure of every editable range in a document.
    void EditableRangeToText()
    {
        auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");
        auto visitor = MakeObject<ExDocumentVisitor::EditableRangeStructurePrinter>();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all of the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc->Accept(visitor);

        std::cout << visitor->GetText() << std::endl;
        TestEditableRangeToText(visitor);
        //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered EditableRange nodes and their children.
    /// </summary>
    class EditableRangeStructurePrinter : public DocumentVisitor
    {
    public:
        EditableRangeStructurePrinter() : mVisitorIsInsideEditableRange(false), mDocTraversalDepth(0)
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
            mVisitorIsInsideEditableRange = false;
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
            // We want to print the contents of runs, but only if they are inside shapes, as they would be in the case of text boxes
            if (mVisitorIsInsideEditableRange)
            {
                IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when an EditableRange node is encountered in the document.
        /// </summary>
        VisitorAction VisitEditableRangeStart(SharedPtr<EditableRangeStart> editableRangeStart) override
        {
            IndentAndAppendLine(String(u"[EditableRange start] ID: ") + editableRangeStart->get_Id() + u" Owner: " +
                                editableRangeStart->get_EditableRange()->get_SingleUser());
            mDocTraversalDepth++;
            mVisitorIsInsideEditableRange = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when the visiting of a EditableRange node is ended.
        /// </summary>
        VisitorAction VisitEditableRangeEnd(SharedPtr<EditableRangeEnd> editableRangeEnd) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[EditableRange end]");
            mVisitorIsInsideEditableRange = false;

            return VisitorAction::Continue;
        }

    private:
        bool mVisitorIsInsideEditableRange;
        int mDocTraversalDepth;
        SharedPtr<System::Text::StringBuilder> mBuilder;

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                mBuilder->Append(u"|  ");
            }

            mBuilder->AppendLine(text);
        }
    };
    //ExEnd

protected:
    void TestEditableRangeToText(SharedPtr<ExDocumentVisitor::EditableRangeStructurePrinter> visitor)
    {
        String visitorText = visitor->GetText();

        ASSERT_TRUE(visitorText.Contains(u"[EditableRange start]"));
        ASSERT_TRUE(visitorText.Contains(u"[EditableRange end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Run]"));
    }

public:
    //ExStart
    //ExFor:DocumentVisitor.VisitFootnoteEnd(Footnote)
    //ExFor:DocumentVisitor.VisitFootnoteStart(Footnote)
    //ExFor:Footnote.Accept(DocumentVisitor)
    //ExSummary:Shows how to print the node structure of every footnote in a document.
    void FootnoteToText()
    {
        auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");
        auto visitor = MakeObject<ExDocumentVisitor::FootnoteStructurePrinter>();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all of the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc->Accept(visitor);

        std::cout << visitor->GetText() << std::endl;
        TestFootnoteToText(visitor);
        //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered Footnote nodes and their children.
    /// </summary>
    class FootnoteStructurePrinter : public DocumentVisitor
    {
    public:
        FootnoteStructurePrinter() : mVisitorIsInsideFootnote(false), mDocTraversalDepth(0)
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
            mVisitorIsInsideFootnote = false;
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        String GetText()
        {
            return mBuilder->ToString();
        }

        /// <summary>
        /// Called when a Footnote node is encountered in the document.
        /// </summary>
        VisitorAction VisitFootnoteStart(SharedPtr<Footnote> footnote) override
        {
            IndentAndAppendLine(String(u"[Footnote start] Type: ") + System::ObjectExt::ToString(footnote->get_FootnoteType()));
            mDocTraversalDepth++;
            mVisitorIsInsideFootnote = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of a Footnote node have been visited.
        /// </summary>
        VisitorAction VisitFootnoteEnd(SharedPtr<Footnote> footnote) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[Footnote end]");
            mVisitorIsInsideFootnote = false;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        VisitorAction VisitRun(SharedPtr<Run> run) override
        {
            if (mVisitorIsInsideFootnote)
            {
                IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
            }

            return VisitorAction::Continue;
        }

    private:
        bool mVisitorIsInsideFootnote;
        int mDocTraversalDepth;
        SharedPtr<System::Text::StringBuilder> mBuilder;

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                mBuilder->Append(u"|  ");
            }

            mBuilder->AppendLine(text);
        }
    };
    //ExEnd

protected:
    void TestFootnoteToText(SharedPtr<ExDocumentVisitor::FootnoteStructurePrinter> visitor)
    {
        String visitorText = visitor->GetText();

        ASSERT_TRUE(visitorText.Contains(u"[Footnote start] Type: Footnote"));
        ASSERT_TRUE(visitorText.Contains(u"[Footnote end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Run]"));
    }

public:
    //ExStart
    //ExFor:DocumentVisitor.VisitOfficeMathEnd(Math.OfficeMath)
    //ExFor:DocumentVisitor.VisitOfficeMathStart(Math.OfficeMath)
    //ExFor:Math.MathObjectType
    //ExFor:Math.OfficeMath.Accept(DocumentVisitor)
    //ExFor:Math.OfficeMath.MathObjectType
    //ExSummary:Shows how to print the node structure of every office math node in a document.
    void OfficeMathToText()
    {
        auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");
        auto visitor = MakeObject<ExDocumentVisitor::OfficeMathStructurePrinter>();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all of the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc->Accept(visitor);

        std::cout << visitor->GetText() << std::endl;
        TestOfficeMathToText(visitor);
        //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered OfficeMath nodes and their children.
    /// </summary>
    class OfficeMathStructurePrinter : public DocumentVisitor
    {
    public:
        OfficeMathStructurePrinter() : mVisitorIsInsideOfficeMath(false), mDocTraversalDepth(0)
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
            mVisitorIsInsideOfficeMath = false;
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
            if (mVisitorIsInsideOfficeMath)
            {
                IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when an OfficeMath node is encountered in the document.
        /// </summary>
        VisitorAction VisitOfficeMathStart(SharedPtr<OfficeMath> officeMath) override
        {
            IndentAndAppendLine(String(u"[OfficeMath start] Math object type: ") + System::ObjectExt::ToString(officeMath->get_MathObjectType()));
            mDocTraversalDepth++;
            mVisitorIsInsideOfficeMath = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of an OfficeMath node have been visited.
        /// </summary>
        VisitorAction VisitOfficeMathEnd(SharedPtr<OfficeMath> officeMath) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[OfficeMath end]");
            mVisitorIsInsideOfficeMath = false;

            return VisitorAction::Continue;
        }

    private:
        bool mVisitorIsInsideOfficeMath;
        int mDocTraversalDepth;
        SharedPtr<System::Text::StringBuilder> mBuilder;

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                mBuilder->Append(u"|  ");
            }

            mBuilder->AppendLine(text);
        }
    };
    //ExEnd

protected:
    void TestOfficeMathToText(SharedPtr<ExDocumentVisitor::OfficeMathStructurePrinter> visitor)
    {
        String visitorText = visitor->GetText();

        ASSERT_TRUE(visitorText.Contains(u"[OfficeMath start] Math object type: OMathPara"));
        ASSERT_TRUE(visitorText.Contains(u"[OfficeMath start] Math object type: OMath"));
        ASSERT_TRUE(visitorText.Contains(u"[OfficeMath start] Math object type: Argument"));
        ASSERT_TRUE(visitorText.Contains(u"[OfficeMath start] Math object type: Supercript"));
        ASSERT_TRUE(visitorText.Contains(u"[OfficeMath start] Math object type: SuperscriptPart"));
        ASSERT_TRUE(visitorText.Contains(u"[OfficeMath start] Math object type: Fraction"));
        ASSERT_TRUE(visitorText.Contains(u"[OfficeMath start] Math object type: Numerator"));
        ASSERT_TRUE(visitorText.Contains(u"[OfficeMath start] Math object type: Denominator"));
        ASSERT_TRUE(visitorText.Contains(u"[OfficeMath end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Run]"));
    }

public:
    //ExStart
    //ExFor:DocumentVisitor.VisitSmartTagEnd(Markup.SmartTag)
    //ExFor:DocumentVisitor.VisitSmartTagStart(Markup.SmartTag)
    //ExSummary:Shows how to print the node structure of every smart tag in a document.
    void SmartTagToText()
    {
        auto doc = MakeObject<Document>(MyDir + u"Smart tags.doc");
        auto visitor = MakeObject<ExDocumentVisitor::SmartTagStructurePrinter>();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all of the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc->Accept(visitor);

        std::cout << visitor->GetText() << std::endl;
        TestSmartTagToText(visitor);
        //ExEnd
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered SmartTag nodes and their children.
    /// </summary>
    class SmartTagStructurePrinter : public DocumentVisitor
    {
    public:
        SmartTagStructurePrinter() : mVisitorIsInsideSmartTag(false), mDocTraversalDepth(0)
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
            mVisitorIsInsideSmartTag = false;
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
            if (mVisitorIsInsideSmartTag)
            {
                IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a SmartTag node is encountered in the document.
        /// </summary>
        VisitorAction VisitSmartTagStart(SharedPtr<SmartTag> smartTag) override
        {
            IndentAndAppendLine(String(u"[SmartTag start] Name: ") + smartTag->get_Element());
            mDocTraversalDepth++;
            mVisitorIsInsideSmartTag = true;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of a SmartTag node have been visited.
        /// </summary>
        VisitorAction VisitSmartTagEnd(SharedPtr<SmartTag> smartTag) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[SmartTag end]");
            mVisitorIsInsideSmartTag = false;

            return VisitorAction::Continue;
        }

    private:
        bool mVisitorIsInsideSmartTag;
        int mDocTraversalDepth;
        SharedPtr<System::Text::StringBuilder> mBuilder;

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                mBuilder->Append(u"|  ");
            }

            mBuilder->AppendLine(text);
        }
    };
    //ExEnd

protected:
    void TestSmartTagToText(SharedPtr<ExDocumentVisitor::SmartTagStructurePrinter> visitor)
    {
        String visitorText = visitor->GetText();

        ASSERT_TRUE(visitorText.Contains(u"[SmartTag start] Name: address"));
        ASSERT_TRUE(visitorText.Contains(u"[SmartTag start] Name: Street"));
        ASSERT_TRUE(visitorText.Contains(u"[SmartTag start] Name: PersonName"));
        ASSERT_TRUE(visitorText.Contains(u"[SmartTag start] Name: title"));
        ASSERT_TRUE(visitorText.Contains(u"[SmartTag start] Name: GivenName"));
        ASSERT_TRUE(visitorText.Contains(u"[SmartTag start] Name: Sn"));
        ASSERT_TRUE(visitorText.Contains(u"[SmartTag start] Name: stockticker"));
        ASSERT_TRUE(visitorText.Contains(u"[SmartTag start] Name: date"));
        ASSERT_TRUE(visitorText.Contains(u"[SmartTag end]"));
        ASSERT_TRUE(visitorText.Contains(u"[Run]"));
    }

public:
    //ExStart
    //ExFor:StructuredDocumentTag.Accept(DocumentVisitor)
    //ExFor:DocumentVisitor.VisitStructuredDocumentTagEnd(Markup.StructuredDocumentTag)
    //ExFor:DocumentVisitor.VisitStructuredDocumentTagStart(Markup.StructuredDocumentTag)
    //ExSummary:Shows how to print the node structure of every structured document tag in a document.
    void StructuredDocumentTagToText()
    {
        auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");
        auto visitor = MakeObject<ExDocumentVisitor::StructuredDocumentTagNodePrinter>();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all of the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc->Accept(visitor);

        std::cout << visitor->GetText() << std::endl;
        TestStructuredDocumentTagToText(visitor);
        //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered StructuredDocumentTag nodes and their children.
    /// </summary>
    class StructuredDocumentTagNodePrinter : public DocumentVisitor
    {
    public:
        StructuredDocumentTagNodePrinter() : mVisitorIsInsideStructuredDocumentTag(false), mDocTraversalDepth(0)
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
            mVisitorIsInsideStructuredDocumentTag = false;
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
            if (mVisitorIsInsideStructuredDocumentTag)
            {
                IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
            }

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a StructuredDocumentTag node is encountered in the document.
        /// </summary>
        VisitorAction VisitStructuredDocumentTagStart(SharedPtr<StructuredDocumentTag> sdt) override
        {
            IndentAndAppendLine(String(u"[StructuredDocumentTag start] Title: ") + sdt->get_Title());
            mDocTraversalDepth++;

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called after all the child nodes of a StructuredDocumentTag node have been visited.
        /// </summary>
        VisitorAction VisitStructuredDocumentTagEnd(SharedPtr<StructuredDocumentTag> sdt) override
        {
            mDocTraversalDepth--;
            IndentAndAppendLine(u"[StructuredDocumentTag end]");

            return VisitorAction::Continue;
        }

    private:
        bool mVisitorIsInsideStructuredDocumentTag;
        int mDocTraversalDepth;
        SharedPtr<System::Text::StringBuilder> mBuilder;

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                mBuilder->Append(u"|  ");
            }

            mBuilder->AppendLine(text);
        }
    };
    //ExEnd

protected:
    void TestStructuredDocumentTagToText(SharedPtr<ExDocumentVisitor::StructuredDocumentTagNodePrinter> visitor)
    {
        String visitorText = visitor->GetText();

        ASSERT_TRUE(visitorText.Contains(u"[StructuredDocumentTag start]"));
        ASSERT_TRUE(visitorText.Contains(u"[StructuredDocumentTag end]"));
    }
};

} // namespace ApiExamples
