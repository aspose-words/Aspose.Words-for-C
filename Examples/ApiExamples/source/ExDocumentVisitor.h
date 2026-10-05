#pragma once

#include <system/text/string_builder.h>
#include <system/string.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeStart.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Math/OfficeMath.h>
#include <Aspose.Words.Cpp/Model/Markup/SmartTag.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTagRangeStart.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTagRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeStart.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/SubDocument.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Math;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExDocumentVisitor : public ApiExampleBase
{
    typedef ExDocumentVisitor ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Traverses a node's tree of child nodes.
    /// Creates a map of this tree in the form of a string.
    /// </summary>
    class DocStructurePrinter : public DocumentVisitor
    {
        typedef DocStructurePrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        DocStructurePrinter();
        
        System::String GetText();
        /// <summary>
        /// Called when a Document node is encountered.
        /// </summary>
        Aspose::Words::VisitorAction VisitDocumentStart(System::SharedPtr<Aspose::Words::Document> doc) override;
        /// <summary>
        /// Called after all the child nodes of a Document node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitDocumentEnd(System::SharedPtr<Aspose::Words::Document> doc) override;
        /// <summary>
        /// Called when a Section node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitSectionStart(System::SharedPtr<Aspose::Words::Section> section) override;
        /// <summary>
        /// Called after all the child nodes of a Section node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitSectionEnd(System::SharedPtr<Aspose::Words::Section> section) override;
        /// <summary>
        /// Called when a Body node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitBodyStart(System::SharedPtr<Aspose::Words::Body> body) override;
        /// <summary>
        /// Called after all the child nodes of a Body node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitBodyEnd(System::SharedPtr<Aspose::Words::Body> body) override;
        /// <summary>
        /// Called when a Paragraph node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitParagraphStart(System::SharedPtr<Aspose::Words::Paragraph> paragraph) override;
        /// <summary>
        /// Called after all the child nodes of a Paragraph node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitParagraphEnd(System::SharedPtr<Aspose::Words::Paragraph> paragraph) override;
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when a SubDocument node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitSubDocument(System::SharedPtr<Aspose::Words::SubDocument> subDocument) override;
        /// <summary>
        /// Called when a SubDocument node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitStructuredDocumentTagRangeStart(System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTagRangeStart> sdtRangeStart) override;
        /// <summary>
        /// Called when a SubDocument node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitStructuredDocumentTagRangeEnd(System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTagRangeEnd> sdtRangeEnd) override;
        
    private:
    
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mAcceptingNodeChildTree;
        
        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered Table nodes and their children.
    /// </summary>
    class TableStructurePrinter : public DocumentVisitor
    {
        typedef TableStructurePrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        TableStructurePrinter();
        
        System::String GetText();
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// Runs that are not within tables are not recorded.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when a Table is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitTableStart(System::SharedPtr<Aspose::Words::Tables::Table> table) override;
        /// <summary>
        /// Called after all the child nodes of a Table node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitTableEnd(System::SharedPtr<Aspose::Words::Tables::Table> table) override;
        /// <summary>
        /// Called when a Row node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRowStart(System::SharedPtr<Aspose::Words::Tables::Row> row) override;
        /// <summary>
        /// Called after all the child nodes of a Row node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitRowEnd(System::SharedPtr<Aspose::Words::Tables::Row> row) override;
        /// <summary>
        /// Called when a Cell node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCellStart(System::SharedPtr<Aspose::Words::Tables::Cell> cell) override;
        /// <summary>
        /// Called after all the child nodes of a Cell node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitCellEnd(System::SharedPtr<Aspose::Words::Tables::Cell> cell) override;
        
    private:
    
        bool mVisitorIsInsideTable;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mVisitedTables;
        
        /// <summary>
        /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is
        /// into the current table's tree of child nodes.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered Comment/CommentRange nodes and their children.
    /// </summary>
    class CommentStructurePrinter : public DocumentVisitor
    {
        typedef CommentStructurePrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        CommentStructurePrinter();
        
        System::String GetText();
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// A Run is only recorded if it is a child of a Comment or CommentRange node.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when a CommentRangeStart node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCommentRangeStart(System::SharedPtr<Aspose::Words::CommentRangeStart> commentRangeStart) override;
        /// <summary>
        /// Called when a CommentRangeEnd node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCommentRangeEnd(System::SharedPtr<Aspose::Words::CommentRangeEnd> commentRangeEnd) override;
        /// <summary>
        /// Called when a Comment node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCommentStart(System::SharedPtr<Aspose::Words::Comment> comment) override;
        /// <summary>
        /// Called after all the child nodes of a Comment node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitCommentEnd(System::SharedPtr<Aspose::Words::Comment> comment) override;
        
    private:
    
        bool mVisitorIsInsideComment;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is
        /// into a comment/comment range's tree of child nodes.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered Field nodes and their children.
    /// </summary>
    class FieldStructurePrinter : public DocumentVisitor
    {
        typedef FieldStructurePrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        FieldStructurePrinter();
        
        System::String GetText();
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFieldStart(System::SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart) override;
        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFieldEnd(System::SharedPtr<Aspose::Words::Fields::FieldEnd> fieldEnd) override;
        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFieldSeparator(System::SharedPtr<Aspose::Words::Fields::FieldSeparator> fieldSeparator) override;
        
    private:
    
        bool mVisitorIsInsideField;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is
        /// into the field's tree of child nodes.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered HeaderFooter nodes and their children.
    /// </summary>
    class HeaderFooterStructurePrinter : public DocumentVisitor
    {
        typedef HeaderFooterStructurePrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        HeaderFooterStructurePrinter();
        
        System::String GetText();
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when a HeaderFooter node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitHeaderFooterStart(System::SharedPtr<Aspose::Words::HeaderFooter> headerFooter) override;
        /// <summary>
        /// Called after all the child nodes of a HeaderFooter node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitHeaderFooterEnd(System::SharedPtr<Aspose::Words::HeaderFooter> headerFooter) override;
        
    private:
    
        bool mVisitorIsInsideHeaderFooter;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered EditableRange nodes and their children.
    /// </summary>
    class EditableRangeStructurePrinter : public DocumentVisitor
    {
        typedef EditableRangeStructurePrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        EditableRangeStructurePrinter();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when an EditableRange node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitEditableRangeStart(System::SharedPtr<Aspose::Words::EditableRangeStart> editableRangeStart) override;
        /// <summary>
        /// Called when the visiting of a EditableRange node is ended.
        /// </summary>
        Aspose::Words::VisitorAction VisitEditableRangeEnd(System::SharedPtr<Aspose::Words::EditableRangeEnd> editableRangeEnd) override;
        
    private:
    
        bool mVisitorIsInsideEditableRange;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered Footnote nodes and their children.
    /// </summary>
    class FootnoteStructurePrinter : public DocumentVisitor
    {
        typedef FootnoteStructurePrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        FootnoteStructurePrinter();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when a Footnote node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFootnoteStart(System::SharedPtr<Aspose::Words::Notes::Footnote> footnote) override;
        /// <summary>
        /// Called after all the child nodes of a Footnote node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitFootnoteEnd(System::SharedPtr<Aspose::Words::Notes::Footnote> footnote) override;
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        
    private:
    
        bool mVisitorIsInsideFootnote;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered OfficeMath nodes and their children.
    /// </summary>
    class OfficeMathStructurePrinter : public DocumentVisitor
    {
        typedef OfficeMathStructurePrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        OfficeMathStructurePrinter();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when an OfficeMath node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitOfficeMathStart(System::SharedPtr<Aspose::Words::Math::OfficeMath> officeMath) override;
        /// <summary>
        /// Called after all the child nodes of an OfficeMath node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitOfficeMathEnd(System::SharedPtr<Aspose::Words::Math::OfficeMath> officeMath) override;
        
    private:
    
        bool mVisitorIsInsideOfficeMath;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered SmartTag nodes and their children.
    /// </summary>
    class SmartTagStructurePrinter : public DocumentVisitor
    {
        typedef SmartTagStructurePrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        SmartTagStructurePrinter();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when a SmartTag node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitSmartTagStart(System::SharedPtr<Aspose::Words::Markup::SmartTag> smartTag) override;
        /// <summary>
        /// Called after all the child nodes of a SmartTag node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitSmartTagEnd(System::SharedPtr<Aspose::Words::Markup::SmartTag> smartTag) override;
        
    private:
    
        bool mVisitorIsInsideSmartTag;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered StructuredDocumentTag nodes and their children.
    /// </summary>
    class StructuredDocumentTagNodePrinter : public DocumentVisitor
    {
        typedef StructuredDocumentTagNodePrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        StructuredDocumentTagNodePrinter();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when a StructuredDocumentTag node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitStructuredDocumentTagStart(System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag> sdt) override;
        /// <summary>
        /// Called after all the child nodes of a StructuredDocumentTag node have been visited.
        /// </summary>
        Aspose::Words::VisitorAction VisitStructuredDocumentTagEnd(System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag> sdt) override;
        
    private:
    
        bool mVisitorIsInsideStructuredDocumentTag;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    
public:

    //ExStart
    //ExFor:Document.Accept(DocumentVisitor)
    //ExFor:Section.Accept(DocumentVisitor)
    //ExFor:SubDocument.Accept(DocumentVisitor)
    //ExFor:CompositeNode.AcceptEnd(DocumentVisitor)
    //ExFor:CompositeNode.AcceptStart(DocumentVisitor)
    //ExFor:Document.AcceptEnd(DocumentVisitor)
    //ExFor:Document.AcceptStart(DocumentVisitor)
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
    //ExFor:DocumentVisitor.VisitStructuredDocumentTagRangeEnd(StructuredDocumentTagRangeEnd)
    //ExFor:DocumentVisitor.VisitStructuredDocumentTagRangeStart(StructuredDocumentTagRangeStart)
    //ExSummary:Shows how to use a document visitor to print a document's node structure.
    void DocStructureToText();
    //ExStart
    //ExFor:Cell.Accept(DocumentVisitor)
    //ExFor:Cell.AcceptStart(DocumentVisitor)
    //ExFor:Cell.AcceptEnd(DocumentVisitor)
    //ExFor:Cell.IsFirstCell
    //ExFor:Cell.IsLastCell
    //ExFor:DocumentVisitor.VisitTableEnd(Table)
    //ExFor:DocumentVisitor.VisitTableStart(Table)
    //ExFor:DocumentVisitor.VisitRowEnd(Row)
    //ExFor:DocumentVisitor.VisitRowStart(Row)
    //ExFor:DocumentVisitor.VisitCellStart(Cell)
    //ExFor:DocumentVisitor.VisitCellEnd(Cell)
    //ExFor:Row.Accept(DocumentVisitor)
    //ExFor:Row.AcceptStart(DocumentVisitor)
    //ExFor:Row.AcceptEnd(DocumentVisitor)
    //ExFor:Row.FirstCell
    //ExFor:Row.GetText
    //ExFor:Row.IsFirstRow
    //ExFor:Row.LastCell
    //ExFor:Row.ParentTable
    //ExSummary:Shows how to print the node structure of every table in a document.
    void TableToText();
    //ExStart
    //ExFor:DocumentVisitor.VisitCommentStart(Comment)
    //ExFor:DocumentVisitor.VisitCommentEnd(Comment)
    //ExFor:DocumentVisitor.VisitCommentRangeEnd(CommentRangeEnd)
    //ExFor:DocumentVisitor.VisitCommentRangeStart(CommentRangeStart)
    //ExSummary:Shows how to print the node structure of every comment and comment range in a document.
    void CommentsToText();
    //ExStart
    //ExFor:DocumentVisitor.VisitFieldStart
    //ExFor:DocumentVisitor.VisitFieldEnd
    //ExFor:DocumentVisitor.VisitFieldSeparator
    //ExSummary:Shows how to print the node structure of every field in a document.
    void FieldToText();
    //ExStart
    //ExFor:DocumentVisitor.VisitHeaderFooterStart(HeaderFooter)
    //ExFor:DocumentVisitor.VisitHeaderFooterEnd(HeaderFooter)
    //ExFor:HeaderFooter.Accept(DocumentVisitor)
    //ExFor:HeaderFooter.AcceptStart(DocumentVisitor)
    //ExFor:HeaderFooter.AcceptEnd(DocumentVisitor)
    //ExFor:HeaderFooterCollection.ToArray
    //ExFor:Run.Accept(DocumentVisitor)
    //ExFor:Run.GetText
    //ExSummary:Shows how to print the node structure of every header and footer in a document.
    void HeaderFooterToText();
    //ExStart
    //ExFor:DocumentVisitor.VisitEditableRangeEnd(EditableRangeEnd)
    //ExFor:DocumentVisitor.VisitEditableRangeStart(EditableRangeStart)
    //ExSummary:Shows how to print the node structure of every editable range in a document.
    void EditableRangeToText();
    //ExStart
    //ExFor:DocumentVisitor.VisitFootnoteEnd(Footnote)
    //ExFor:DocumentVisitor.VisitFootnoteStart(Footnote)
    //ExFor:Footnote.Accept(DocumentVisitor)
    //ExFor:Footnote.AcceptStart(DocumentVisitor)
    //ExFor:Footnote.AcceptEnd(DocumentVisitor)
    //ExSummary:Shows how to print the node structure of every footnote in a document.
    void FootnoteToText();
    //ExStart
    //ExFor:DocumentVisitor.VisitOfficeMathEnd(OfficeMath)
    //ExFor:DocumentVisitor.VisitOfficeMathStart(OfficeMath)
    //ExFor:MathObjectType
    //ExFor:OfficeMath.Accept(DocumentVisitor)
    //ExFor:OfficeMath.AcceptStart(DocumentVisitor)
    //ExFor:OfficeMath.AcceptEnd(DocumentVisitor)
    //ExFor:OfficeMath.MathObjectType
    //ExSummary:Shows how to print the node structure of every office math node in a document.
    void OfficeMathToText();
    //ExStart
    //ExFor:DocumentVisitor.VisitSmartTagEnd(SmartTag)
    //ExFor:DocumentVisitor.VisitSmartTagStart(SmartTag)
    //ExSummary:Shows how to print the node structure of every smart tag in a document.
    void SmartTagToText();
    //ExStart
    //ExFor:StructuredDocumentTag.Accept(DocumentVisitor)
    //ExFor:StructuredDocumentTag.AcceptStart(DocumentVisitor)
    //ExFor:StructuredDocumentTag.AcceptEnd(DocumentVisitor)
    //ExFor:DocumentVisitor.VisitStructuredDocumentTagEnd(StructuredDocumentTag)
    //ExFor:DocumentVisitor.VisitStructuredDocumentTagStart(StructuredDocumentTag)
    //ExSummary:Shows how to print the node structure of every structured document tag in a document.
    void StructuredDocumentTagToText();
    
protected:

    //ExEnd
    static void TestDocStructureToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::DocStructurePrinter> visitor);
    //ExEnd
    static void TestTableToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::TableStructurePrinter> visitor);
    //ExEnd
    static void TestCommentsToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::CommentStructurePrinter> visitor);
    //ExEnd
    static void TestFieldToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::FieldStructurePrinter> visitor);
    //ExEnd
    static void TestHeaderFooterToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::HeaderFooterStructurePrinter> visitor);
    //ExEnd
    static void TestEditableRangeToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::EditableRangeStructurePrinter> visitor);
    //ExEnd
    static void TestFootnoteToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::FootnoteStructurePrinter> visitor);
    //ExEnd
    static void TestOfficeMathToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::OfficeMathStructurePrinter> visitor);
    //ExEnd
    static void TestSmartTagToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::SmartTagStructurePrinter> visitor);
    //ExEnd
    static void TestStructuredDocumentTagToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::StructuredDocumentTagNodePrinter> visitor);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


