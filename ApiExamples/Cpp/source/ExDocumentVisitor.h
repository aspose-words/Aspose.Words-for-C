#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { enum class VisitorAction; } }
namespace Aspose { namespace Words { class Document; } }
namespace Aspose { namespace Words { class Section; } }
namespace Aspose { namespace Words { class Body; } }
namespace Aspose { namespace Words { class Paragraph; } }
namespace Aspose { namespace Words { class Run; } }
namespace Aspose { namespace Words { class SubDocument; } }
namespace Aspose { namespace Words { namespace Tables { class Table; } } }
namespace Aspose { namespace Words { namespace Tables { class Row; } } }
namespace Aspose { namespace Words { namespace Tables { class Cell; } } }
namespace Aspose { namespace Words { class CommentRangeStart; } }
namespace Aspose { namespace Words { class CommentRangeEnd; } }
namespace Aspose { namespace Words { class Comment; } }
namespace Aspose { namespace Words { namespace Fields { class FieldStart; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldEnd; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldSeparator; } } }
namespace Aspose { namespace Words { class HeaderFooter; } }
namespace Aspose { namespace Words { class EditableRangeStart; } }
namespace Aspose { namespace Words { class EditableRangeEnd; } }
namespace Aspose { namespace Words { class Footnote; } }
namespace Aspose { namespace Words { namespace Math { class OfficeMath; } } }
namespace Aspose { namespace Words { namespace Markup { class SmartTag; } } }
namespace Aspose { namespace Words { namespace Markup { class StructuredDocumentTag; } } }

namespace ApiExamples {

class ExDocumentVisitor : public ApiExampleBase
{
public:

    /// <summary>
    /// This Visitor implementation prints information about sections, bodies, paragraphs and runs encountered in the document.
    /// </summary>
    class DocStructurePrinter : public Aspose::Words::DocumentVisitor
    {
        typedef DocStructurePrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        DocStructurePrinter();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when a Document node is encountered.
        /// </summary>
        Aspose::Words::VisitorAction VisitDocumentStart(System::SharedPtr<Aspose::Words::Document> doc) override;
        /// <summary>
        /// Called when the visiting of a Document is ended.
        /// </summary>
        Aspose::Words::VisitorAction VisitDocumentEnd(System::SharedPtr<Aspose::Words::Document> doc) override;
        /// <summary>
        /// Called when a Section node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitSectionStart(System::SharedPtr<Aspose::Words::Section> section) override;
        /// <summary>
        /// Called when the visiting of a Section node is ended.
        /// </summary>
        Aspose::Words::VisitorAction VisitSectionEnd(System::SharedPtr<Aspose::Words::Section> section) override;
        /// <summary>
        /// Called when a Body node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitBodyStart(System::SharedPtr<Aspose::Words::Body> body) override;
        /// <summary>
        /// Called when the visiting of a Body node is ended.
        /// </summary>
        Aspose::Words::VisitorAction VisitBodyEnd(System::SharedPtr<Aspose::Words::Body> body) override;
        /// <summary>
        /// Called when a Paragraph node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitParagraphStart(System::SharedPtr<Aspose::Words::Paragraph> paragraph) override;
        /// <summary>
        /// Called when the visiting of a Paragraph node is ended.
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
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// This Visitor implementation prints information about and contents of all tables encountered in the document.
    /// </summary>
    class TableInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef TableInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        TableInfoPrinter();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when a Table is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitTableStart(System::SharedPtr<Aspose::Words::Tables::Table> table) override;
        /// <summary>
        /// Called when the visiting of a Table node is ended.
        /// </summary>
        Aspose::Words::VisitorAction VisitTableEnd(System::SharedPtr<Aspose::Words::Tables::Table> table) override;
        /// <summary>
        /// Called when a Row node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRowStart(System::SharedPtr<Aspose::Words::Tables::Row> row) override;
        /// <summary>
        /// Called when the visiting of a Row node is ended.
        /// </summary>
        Aspose::Words::VisitorAction VisitRowEnd(System::SharedPtr<Aspose::Words::Tables::Row> row) override;
        /// <summary>
        /// Called when a Cell node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCellStart(System::SharedPtr<Aspose::Words::Tables::Cell> cell) override;
        /// <summary>
        /// Called when the visiting of a Cell node is ended in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCellEnd(System::SharedPtr<Aspose::Words::Tables::Cell> cell) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        bool mVisitorIsInsideTable;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// This Visitor implementation prints information about and contents of comments and comment ranges encountered in the document.
    /// </summary>
    class CommentInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef CommentInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        CommentInfoPrinter();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when a Run node is encountered in the document.
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
        /// Called when the visiting of a Comment node is ended in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCommentEnd(System::SharedPtr<Aspose::Words::Comment> comment) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        bool mVisitorIsInsideComment;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// This Visitor implementation prints information about fields encountered in the document.
    /// </summary>
    class FieldInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef FieldInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        FieldInfoPrinter();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
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
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        bool mVisitorIsInsideField;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// This Visitor implementation prints information about HeaderFooter nodes encountered in the document.
    /// </summary>
    class HeaderFooterInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef HeaderFooterInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        HeaderFooterInfoPrinter();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
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
        /// Called when the visiting of a HeaderFooter node is ended.
        /// </summary>
        Aspose::Words::VisitorAction VisitHeaderFooterEnd(System::SharedPtr<Aspose::Words::HeaderFooter> headerFooter) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        bool mVisitorIsInsideHeaderFooter;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    /// <summary>
    /// This Visitor implementation prints information about editable ranges encountered in the document.
    /// </summary>
    class EditableRangeInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef EditableRangeInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        EditableRangeInfoPrinter();
        
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
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
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
    /// This Visitor implementation prints information about footnotes encountered in the document.
    /// </summary>
    class FootnoteInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef FootnoteInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        FootnoteInfoPrinter();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when a Footnote node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFootnoteStart(System::SharedPtr<Aspose::Words::Footnote> footnote) override;
        /// <summary>
        /// Called when the visiting of a Footnote node is ended.
        /// </summary>
        Aspose::Words::VisitorAction VisitFootnoteEnd(System::SharedPtr<Aspose::Words::Footnote> footnote) override;
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
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
    /// This Visitor implementation prints information about office math objects encountered in the document.
    /// </summary>
    class OfficeMathInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef OfficeMathInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        OfficeMathInfoPrinter();
        
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
        /// Called when the visiting of a OfficeMath node is ended.
        /// </summary>
        Aspose::Words::VisitorAction VisitOfficeMathEnd(System::SharedPtr<Aspose::Words::Math::OfficeMath> officeMath) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
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
    /// This Visitor implementation prints information about smart tags encountered in the document.
    /// </summary>
    class SmartTagInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef SmartTagInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        SmartTagInfoPrinter();
        
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
        /// Called when the visiting of a SmartTag node is ended.
        /// </summary>
        Aspose::Words::VisitorAction VisitSmartTagEnd(System::SharedPtr<Aspose::Words::Markup::SmartTag> smartTag) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
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
    /// This Visitor implementation prints information about structured document tags encountered in the document.
    /// </summary>
    class StructuredDocumentTagInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef StructuredDocumentTagInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        StructuredDocumentTagInfoPrinter();
        
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
        /// Called when the visiting of a StructuredDocumentTag node is ended.
        /// </summary>
        Aspose::Words::VisitorAction VisitStructuredDocumentTagEnd(System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag> sdt) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
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

    void DocStructureToText();
    void TableToText();
    void CommentsToText();
    void FieldToText();
    void HeaderFooterToText();
    void EditableRangeToText();
    void FootnoteToText();
    void OfficeMathToText();
    void SmartTagToText();
    void StructuredDocumentTagToText();
    
protected:

    void TestDocStructureToText(System::SharedPtr<ExDocumentVisitor::DocStructurePrinter> visitor);
    void TestTableToText(System::SharedPtr<ExDocumentVisitor::TableInfoPrinter> visitor);
    void TestCommentsToText(System::SharedPtr<ExDocumentVisitor::CommentInfoPrinter> visitor);
    void TestFieldToText(System::SharedPtr<ExDocumentVisitor::FieldInfoPrinter> visitor);
    void TestHeaderFooterToText(System::SharedPtr<ExDocumentVisitor::HeaderFooterInfoPrinter> visitor);
    void TestEditableRangeToText(System::SharedPtr<ExDocumentVisitor::EditableRangeInfoPrinter> visitor);
    void TestFootnoteToText(System::SharedPtr<ExDocumentVisitor::FootnoteInfoPrinter> visitor);
    void TestOfficeMathToText(System::SharedPtr<ExDocumentVisitor::OfficeMathInfoPrinter> visitor);
    void TestSmartTagToText(System::SharedPtr<ExDocumentVisitor::SmartTagInfoPrinter> visitor);
    void TestStructuredDocumentTagToText(System::SharedPtr<ExDocumentVisitor::StructuredDocumentTagInfoPrinter> visitor);
    
};

} // namespace ApiExamples


