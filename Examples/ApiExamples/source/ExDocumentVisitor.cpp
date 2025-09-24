// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentVisitor.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/date_time.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Math/MathObjectType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRange.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>


using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Math;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1860046044u, ::Aspose::Words::ApiExamples::ExDocumentVisitor::DocStructurePrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::DocStructurePrinter::DocStructurePrinter() : mDocTraversalDepth(0)
{
    mAcceptingNodeChildTree = System::MakeObject<System::Text::StringBuilder>();
}

System::String ExDocumentVisitor::DocStructurePrinter::GetText()
{
    return mAcceptingNodeChildTree->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitDocumentStart(System::SharedPtr<Aspose::Words::Document> doc)
{
    int32_t childNodeCount = doc->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count();
    
    IndentAndAppendLine(System::String(u"[Document start] Child nodes: ") + childNodeCount);
    mDocTraversalDepth++;
    
    // Allow the visitor to continue visiting other nodes.
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitDocumentEnd(System::SharedPtr<Aspose::Words::Document> doc)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Document end]");
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitSectionStart(System::SharedPtr<Aspose::Words::Section> section)
{
    // Get the index of our section within the document.
    System::SharedPtr<Aspose::Words::NodeCollection> docSections = section->get_Document()->GetChildNodes(Aspose::Words::NodeType::Section, false);
    int32_t sectionIndex = docSections->IndexOf(section);
    
    IndentAndAppendLine(System::String(u"[Section start] Section index: ") + sectionIndex);
    mDocTraversalDepth++;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitSectionEnd(System::SharedPtr<Aspose::Words::Section> section)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Section end]");
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitBodyStart(System::SharedPtr<Aspose::Words::Body> body)
{
    int32_t paragraphCount = body->get_Paragraphs()->get_Count();
    IndentAndAppendLine(System::String(u"[Body start] Paragraphs: ") + paragraphCount);
    mDocTraversalDepth++;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitBodyEnd(System::SharedPtr<Aspose::Words::Body> body)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Body end]");
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitParagraphStart(System::SharedPtr<Aspose::Words::Paragraph> paragraph)
{
    IndentAndAppendLine(u"[Paragraph start]");
    mDocTraversalDepth++;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitParagraphEnd(System::SharedPtr<Aspose::Words::Paragraph> paragraph)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Paragraph end]");
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    IndentAndAppendLine(System::String(u"[Run] \"") + run->GetText() + u"\"");
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitSubDocument(System::SharedPtr<Aspose::Words::SubDocument> subDocument)
{
    IndentAndAppendLine(u"[SubDocument]");
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitStructuredDocumentTagRangeStart(System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTagRangeStart> sdtRangeStart)
{
    IndentAndAppendLine(u"[SdtRangeStart]");
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitStructuredDocumentTagRangeEnd(System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTagRangeEnd> sdtRangeEnd)
{
    IndentAndAppendLine(u"[SdtRangeEnd]");
    
    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::DocStructurePrinter::IndentAndAppendLine(System::String text)
{
    for (int32_t i = 0; i < mDocTraversalDepth; i++)
    {
        mAcceptingNodeChildTree->Append(u"|  ");
    }
    
    mAcceptingNodeChildTree->AppendLine(text);
}

RTTI_INFO_IMPL_HASH(2491243378u, ::Aspose::Words::ApiExamples::ExDocumentVisitor::TableStructurePrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::TableStructurePrinter::TableStructurePrinter() : mVisitorIsInsideTable(false)
    , mDocTraversalDepth(0)
{
    mVisitedTables = System::MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideTable = false;
}

System::String ExDocumentVisitor::TableStructurePrinter::GetText()
{
    return mVisitedTables->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableStructurePrinter::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideTable)
    {
        IndentAndAppendLine(System::String(u"[Run] \"") + run->GetText() + u"\"");
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableStructurePrinter::VisitTableStart(System::SharedPtr<Aspose::Words::Tables::Table> table)
{
    int32_t rows = 0;
    int32_t columns = 0;
    
    if (table->get_Rows()->get_Count() > 0)
    {
        rows = table->get_Rows()->get_Count();
        columns = table->get_FirstRow()->get_Count();
    }
    
    IndentAndAppendLine(System::String(u"[Table start] Size: ") + rows + u"x" + columns);
    mDocTraversalDepth++;
    mVisitorIsInsideTable = true;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableStructurePrinter::VisitTableEnd(System::SharedPtr<Aspose::Words::Tables::Table> table)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Table end]");
    mVisitorIsInsideTable = false;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableStructurePrinter::VisitRowStart(System::SharedPtr<Aspose::Words::Tables::Row> row)
{
    System::String rowContents = row->GetText().TrimEnd(System::MakeArray<char16_t>({u'\x0007', u' '})).Replace(u"\u0007", u", ");
    int32_t rowWidth = row->IndexOf(row->get_LastCell()) + 1;
    int32_t rowIndex = row->get_ParentTable()->IndexOf(row);
    System::String rowStatusInTable = row->get_IsFirstRow() && row->get_IsLastRow() ? u"only" : row->get_IsFirstRow() ? u"first" : row->get_IsLastRow() ? System::String(u"last") : System::String(u"");
    if (rowStatusInTable != u"")
    {
        rowStatusInTable = System::String::Format(u", the {0} row in this table,", rowStatusInTable);
    }
    
    IndentAndAppendLine(System::String::Format(u"[Row start] Row #{0}{1} width {2}, \"{3}\"", ++rowIndex, rowStatusInTable, rowWidth, rowContents));
    mDocTraversalDepth++;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableStructurePrinter::VisitRowEnd(System::SharedPtr<Aspose::Words::Tables::Row> row)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Row end]");
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableStructurePrinter::VisitCellStart(System::SharedPtr<Aspose::Words::Tables::Cell> cell)
{
    System::SharedPtr<Aspose::Words::Tables::Row> row = cell->get_ParentRow();
    System::SharedPtr<Aspose::Words::Tables::Table> table = row->get_ParentTable();
    System::String cellStatusInRow = cell->get_IsFirstCell() && cell->get_IsLastCell() ? u"only" : cell->get_IsFirstCell() ? u"first" : cell->get_IsLastCell() ? System::String(u"last") : System::String(u"");
    if (cellStatusInRow != u"")
    {
        cellStatusInRow = System::String::Format(u", the {0} cell in this row", cellStatusInRow);
    }
    
    IndentAndAppendLine(System::String::Format(u"[Cell start] Row {0}, Col {1}{2}", table->IndexOf(row) + 1, row->IndexOf(cell) + 1, cellStatusInRow));
    mDocTraversalDepth++;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableStructurePrinter::VisitCellEnd(System::SharedPtr<Aspose::Words::Tables::Cell> cell)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Cell end]");
    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::TableStructurePrinter::IndentAndAppendLine(System::String text)
{
    for (int32_t i = 0; i < mDocTraversalDepth; i++)
    {
        mVisitedTables->Append(u"|  ");
    }
    
    mVisitedTables->AppendLine(text);
}

RTTI_INFO_IMPL_HASH(1329380899u, ::Aspose::Words::ApiExamples::ExDocumentVisitor::CommentStructurePrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::CommentStructurePrinter::CommentStructurePrinter() : mVisitorIsInsideComment(false)
    , mDocTraversalDepth(0)
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideComment = false;
}

System::String ExDocumentVisitor::CommentStructurePrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::CommentStructurePrinter::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideComment)
    {
        IndentAndAppendLine(System::String(u"[Run] \"") + run->GetText() + u"\"");
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::CommentStructurePrinter::VisitCommentRangeStart(System::SharedPtr<Aspose::Words::CommentRangeStart> commentRangeStart)
{
    IndentAndAppendLine(System::String(u"[Comment range start] ID: ") + commentRangeStart->get_Id());
    mDocTraversalDepth++;
    mVisitorIsInsideComment = true;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::CommentStructurePrinter::VisitCommentRangeEnd(System::SharedPtr<Aspose::Words::CommentRangeEnd> commentRangeEnd)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Comment range end]");
    mVisitorIsInsideComment = false;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::CommentStructurePrinter::VisitCommentStart(System::SharedPtr<Aspose::Words::Comment> comment)
{
    IndentAndAppendLine(System::String::Format(u"[Comment start] For comment range ID {0}, By {1} on {2}", comment->get_Id(), comment->get_Author(), comment->get_DateTime()));
    mDocTraversalDepth++;
    mVisitorIsInsideComment = true;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::CommentStructurePrinter::VisitCommentEnd(System::SharedPtr<Aspose::Words::Comment> comment)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Comment end]");
    mVisitorIsInsideComment = false;
    
    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::CommentStructurePrinter::IndentAndAppendLine(System::String text)
{
    for (int32_t i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }
    
    mBuilder->AppendLine(text);
}

RTTI_INFO_IMPL_HASH(1604398558u, ::Aspose::Words::ApiExamples::ExDocumentVisitor::FieldStructurePrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::FieldStructurePrinter::FieldStructurePrinter() : mVisitorIsInsideField(false)
    , mDocTraversalDepth(0)
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideField = false;
}

System::String ExDocumentVisitor::FieldStructurePrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::FieldStructurePrinter::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideField)
    {
        IndentAndAppendLine(System::String(u"[Run] \"") + run->GetText() + u"\"");
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::FieldStructurePrinter::VisitFieldStart(System::SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart)
{
    IndentAndAppendLine(System::String(u"[Field start] FieldType: ") + System::ObjectExt::ToString(fieldStart->get_FieldType()));
    mDocTraversalDepth++;
    mVisitorIsInsideField = true;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::FieldStructurePrinter::VisitFieldEnd(System::SharedPtr<Aspose::Words::Fields::FieldEnd> fieldEnd)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Field end]");
    mVisitorIsInsideField = false;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::FieldStructurePrinter::VisitFieldSeparator(System::SharedPtr<Aspose::Words::Fields::FieldSeparator> fieldSeparator)
{
    IndentAndAppendLine(u"[FieldSeparator]");
    
    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::FieldStructurePrinter::IndentAndAppendLine(System::String text)
{
    for (int32_t i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }
    
    mBuilder->AppendLine(text);
}

RTTI_INFO_IMPL_HASH(3446194194u, ::Aspose::Words::ApiExamples::ExDocumentVisitor::HeaderFooterStructurePrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::HeaderFooterStructurePrinter::HeaderFooterStructurePrinter() : mVisitorIsInsideHeaderFooter(false)
    , mDocTraversalDepth(0)
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideHeaderFooter = false;
}

System::String ExDocumentVisitor::HeaderFooterStructurePrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::HeaderFooterStructurePrinter::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideHeaderFooter)
    {
        IndentAndAppendLine(System::String(u"[Run] \"") + run->GetText() + u"\"");
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::HeaderFooterStructurePrinter::VisitHeaderFooterStart(System::SharedPtr<Aspose::Words::HeaderFooter> headerFooter)
{
    IndentAndAppendLine(System::String(u"[HeaderFooter start] HeaderFooterType: ") + System::ObjectExt::ToString(headerFooter->get_HeaderFooterType()));
    mDocTraversalDepth++;
    mVisitorIsInsideHeaderFooter = true;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::HeaderFooterStructurePrinter::VisitHeaderFooterEnd(System::SharedPtr<Aspose::Words::HeaderFooter> headerFooter)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[HeaderFooter end]");
    mVisitorIsInsideHeaderFooter = false;
    
    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::HeaderFooterStructurePrinter::IndentAndAppendLine(System::String text)
{
    for (int32_t i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }
    
    mBuilder->AppendLine(text);
}

RTTI_INFO_IMPL_HASH(1581082717u, ::Aspose::Words::ApiExamples::ExDocumentVisitor::EditableRangeStructurePrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::EditableRangeStructurePrinter::EditableRangeStructurePrinter()
    : mVisitorIsInsideEditableRange(false), mDocTraversalDepth(0)
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideEditableRange = false;
}

System::String ExDocumentVisitor::EditableRangeStructurePrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::EditableRangeStructurePrinter::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    // We want to print the contents of runs, but only if they are inside shapes, as they would be in the case of text boxes
    if (mVisitorIsInsideEditableRange)
    {
        IndentAndAppendLine(System::String(u"[Run] \"") + run->GetText() + u"\"");
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::EditableRangeStructurePrinter::VisitEditableRangeStart(System::SharedPtr<Aspose::Words::EditableRangeStart> editableRangeStart)
{
    IndentAndAppendLine(System::String(u"[EditableRange start] ID: ") + editableRangeStart->get_Id() + u" Owner: " + editableRangeStart->get_EditableRange()->get_SingleUser());
    mDocTraversalDepth++;
    mVisitorIsInsideEditableRange = true;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::EditableRangeStructurePrinter::VisitEditableRangeEnd(System::SharedPtr<Aspose::Words::EditableRangeEnd> editableRangeEnd)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[EditableRange end]");
    mVisitorIsInsideEditableRange = false;
    
    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::EditableRangeStructurePrinter::IndentAndAppendLine(System::String text)
{
    for (int32_t i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }
    
    mBuilder->AppendLine(text);
}

RTTI_INFO_IMPL_HASH(3913636106u, ::Aspose::Words::ApiExamples::ExDocumentVisitor::FootnoteStructurePrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::FootnoteStructurePrinter::FootnoteStructurePrinter() : mVisitorIsInsideFootnote(false)
    , mDocTraversalDepth(0)
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideFootnote = false;
}

System::String ExDocumentVisitor::FootnoteStructurePrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::FootnoteStructurePrinter::VisitFootnoteStart(System::SharedPtr<Aspose::Words::Notes::Footnote> footnote)
{
    IndentAndAppendLine(System::String(u"[Footnote start] Type: ") + System::ObjectExt::ToString(footnote->get_FootnoteType()));
    mDocTraversalDepth++;
    mVisitorIsInsideFootnote = true;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::FootnoteStructurePrinter::VisitFootnoteEnd(System::SharedPtr<Aspose::Words::Notes::Footnote> footnote)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Footnote end]");
    mVisitorIsInsideFootnote = false;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::FootnoteStructurePrinter::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideFootnote)
    {
        IndentAndAppendLine(System::String(u"[Run] \"") + run->GetText() + u"\"");
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::FootnoteStructurePrinter::IndentAndAppendLine(System::String text)
{
    for (int32_t i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }
    
    mBuilder->AppendLine(text);
}

RTTI_INFO_IMPL_HASH(1473641614u, ::Aspose::Words::ApiExamples::ExDocumentVisitor::OfficeMathStructurePrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::OfficeMathStructurePrinter::OfficeMathStructurePrinter() : mVisitorIsInsideOfficeMath(false)
    , mDocTraversalDepth(0)
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideOfficeMath = false;
}

System::String ExDocumentVisitor::OfficeMathStructurePrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::OfficeMathStructurePrinter::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideOfficeMath)
    {
        IndentAndAppendLine(System::String(u"[Run] \"") + run->GetText() + u"\"");
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::OfficeMathStructurePrinter::VisitOfficeMathStart(System::SharedPtr<Aspose::Words::Math::OfficeMath> officeMath)
{
    IndentAndAppendLine(System::String(u"[OfficeMath start] Math object type: ") + System::ObjectExt::ToString(officeMath->get_MathObjectType()));
    mDocTraversalDepth++;
    mVisitorIsInsideOfficeMath = true;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::OfficeMathStructurePrinter::VisitOfficeMathEnd(System::SharedPtr<Aspose::Words::Math::OfficeMath> officeMath)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[OfficeMath end]");
    mVisitorIsInsideOfficeMath = false;
    
    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::OfficeMathStructurePrinter::IndentAndAppendLine(System::String text)
{
    for (int32_t i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }
    
    mBuilder->AppendLine(text);
}

RTTI_INFO_IMPL_HASH(1392209019u, ::Aspose::Words::ApiExamples::ExDocumentVisitor::SmartTagStructurePrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::SmartTagStructurePrinter::SmartTagStructurePrinter() : mVisitorIsInsideSmartTag(false)
    , mDocTraversalDepth(0)
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideSmartTag = false;
}

System::String ExDocumentVisitor::SmartTagStructurePrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::SmartTagStructurePrinter::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideSmartTag)
    {
        IndentAndAppendLine(System::String(u"[Run] \"") + run->GetText() + u"\"");
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::SmartTagStructurePrinter::VisitSmartTagStart(System::SharedPtr<Aspose::Words::Markup::SmartTag> smartTag)
{
    IndentAndAppendLine(System::String(u"[SmartTag start] Name: ") + smartTag->get_Element());
    mDocTraversalDepth++;
    mVisitorIsInsideSmartTag = true;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::SmartTagStructurePrinter::VisitSmartTagEnd(System::SharedPtr<Aspose::Words::Markup::SmartTag> smartTag)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[SmartTag end]");
    mVisitorIsInsideSmartTag = false;
    
    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::SmartTagStructurePrinter::IndentAndAppendLine(System::String text)
{
    for (int32_t i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }
    
    mBuilder->AppendLine(text);
}

RTTI_INFO_IMPL_HASH(2328101517u, ::Aspose::Words::ApiExamples::ExDocumentVisitor::StructuredDocumentTagNodePrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::StructuredDocumentTagNodePrinter::StructuredDocumentTagNodePrinter()
    : mVisitorIsInsideStructuredDocumentTag(false), mDocTraversalDepth(0)
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideStructuredDocumentTag = false;
}

System::String ExDocumentVisitor::StructuredDocumentTagNodePrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::StructuredDocumentTagNodePrinter::VisitRun(System::SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideStructuredDocumentTag)
    {
        IndentAndAppendLine(System::String(u"[Run] \"") + run->GetText() + u"\"");
    }
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::StructuredDocumentTagNodePrinter::VisitStructuredDocumentTagStart(System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag> sdt)
{
    IndentAndAppendLine(System::String(u"[StructuredDocumentTag start] Title: ") + sdt->get_Title());
    mDocTraversalDepth++;
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::StructuredDocumentTagNodePrinter::VisitStructuredDocumentTagEnd(System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag> sdt)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[StructuredDocumentTag end]");
    
    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::StructuredDocumentTagNodePrinter::IndentAndAppendLine(System::String text)
{
    for (int32_t i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }
    
    mBuilder->AppendLine(text);
}


RTTI_INFO_IMPL_HASH(1841784003u, ::Aspose::Words::ApiExamples::ExDocumentVisitor, ThisTypeBaseTypesInfo);

void ExDocumentVisitor::TestDocStructureToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::DocStructurePrinter> visitor)
{
    System::String visitorText = visitor->GetText();
    
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

void ExDocumentVisitor::TestTableToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::TableStructurePrinter> visitor)
{
    System::String visitorText = visitor->GetText();
    
    ASSERT_TRUE(visitorText.Contains(u"[Table start]"));
    ASSERT_TRUE(visitorText.Contains(u"[Table end]"));
    ASSERT_TRUE(visitorText.Contains(u"[Row start]"));
    ASSERT_TRUE(visitorText.Contains(u"[Row end]"));
    ASSERT_TRUE(visitorText.Contains(u"[Cell start]"));
    ASSERT_TRUE(visitorText.Contains(u"[Cell end]"));
    ASSERT_TRUE(visitorText.Contains(u"[Run]"));
}

void ExDocumentVisitor::TestCommentsToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::CommentStructurePrinter> visitor)
{
    System::String visitorText = visitor->GetText();
    
    ASSERT_TRUE(visitorText.Contains(u"[Comment range start]"));
    ASSERT_TRUE(visitorText.Contains(u"[Comment range end]"));
    ASSERT_TRUE(visitorText.Contains(u"[Comment start]"));
    ASSERT_TRUE(visitorText.Contains(u"[Comment end]"));
    ASSERT_TRUE(visitorText.Contains(u"[Run]"));
}

void ExDocumentVisitor::TestFieldToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::FieldStructurePrinter> visitor)
{
    System::String visitorText = visitor->GetText();
    
    ASSERT_TRUE(visitorText.Contains(u"[Field start]"));
    ASSERT_TRUE(visitorText.Contains(u"[Field end]"));
    ASSERT_TRUE(visitorText.Contains(u"[FieldSeparator]"));
    ASSERT_TRUE(visitorText.Contains(u"[Run]"));
}

void ExDocumentVisitor::TestHeaderFooterToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::HeaderFooterStructurePrinter> visitor)
{
    System::String visitorText = visitor->GetText();
    
    ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter start] HeaderFooterType: HeaderPrimary"));
    ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter end]"));
    ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter start] HeaderFooterType: HeaderFirst"));
    ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter start] HeaderFooterType: HeaderEven"));
    ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter start] HeaderFooterType: FooterPrimary"));
    ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter start] HeaderFooterType: FooterFirst"));
    ASSERT_TRUE(visitorText.Contains(u"[HeaderFooter start] HeaderFooterType: FooterEven"));
    ASSERT_TRUE(visitorText.Contains(u"[Run]"));
}

void ExDocumentVisitor::TestEditableRangeToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::EditableRangeStructurePrinter> visitor)
{
    System::String visitorText = visitor->GetText();
    
    ASSERT_TRUE(visitorText.Contains(u"[EditableRange start]"));
    ASSERT_TRUE(visitorText.Contains(u"[EditableRange end]"));
    ASSERT_TRUE(visitorText.Contains(u"[Run]"));
}

void ExDocumentVisitor::TestFootnoteToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::FootnoteStructurePrinter> visitor)
{
    System::String visitorText = visitor->GetText();
    
    ASSERT_TRUE(visitorText.Contains(u"[Footnote start] Type: Footnote"));
    ASSERT_TRUE(visitorText.Contains(u"[Footnote end]"));
    ASSERT_TRUE(visitorText.Contains(u"[Run]"));
}

void ExDocumentVisitor::TestOfficeMathToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::OfficeMathStructurePrinter> visitor)
{
    System::String visitorText = visitor->GetText();
    
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

void ExDocumentVisitor::TestSmartTagToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::SmartTagStructurePrinter> visitor)
{
    System::String visitorText = visitor->GetText();
    
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

void ExDocumentVisitor::TestStructuredDocumentTagToText(System::SharedPtr<Aspose::Words::ApiExamples::ExDocumentVisitor::StructuredDocumentTagNodePrinter> visitor)
{
    System::String visitorText = visitor->GetText();
    
    ASSERT_TRUE(visitorText.Contains(u"[StructuredDocumentTag start]"));
    ASSERT_TRUE(visitorText.Contains(u"[StructuredDocumentTag end]"));
}


namespace gtest_test
{

class ExDocumentVisitor : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExDocumentVisitor> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExDocumentVisitor>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExDocumentVisitor> ExDocumentVisitor::s_instance;

} // namespace gtest_test

void ExDocumentVisitor::DocStructureToText()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DocumentVisitor-compatible features.docx");
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExDocumentVisitor::DocStructurePrinter>();
    
    // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
    // and then traverses all the node's children in a depth-first manner.
    // The visitor can read and modify each visited node.
    doc->Accept(visitor);
    
    std::cout << visitor->GetText() << std::endl;
    TestDocStructureToText(visitor);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocumentVisitor, DocStructureToText)
{
    s_instance->DocStructureToText();
}

} // namespace gtest_test

void ExDocumentVisitor::TableToText()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DocumentVisitor-compatible features.docx");
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExDocumentVisitor::TableStructurePrinter>();
    
    // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
    // and then traverses all the node's children in a depth-first manner.
    // The visitor can read and modify each visited node.
    doc->Accept(visitor);
    
    std::cout << visitor->GetText() << std::endl;
    TestTableToText(visitor);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocumentVisitor, TableToText)
{
    s_instance->TableToText();
}

} // namespace gtest_test

void ExDocumentVisitor::CommentsToText()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DocumentVisitor-compatible features.docx");
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExDocumentVisitor::CommentStructurePrinter>();
    
    // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
    // and then traverses all the node's children in a depth-first manner.
    // The visitor can read and modify each visited node.
    doc->Accept(visitor);
    
    std::cout << visitor->GetText() << std::endl;
    TestCommentsToText(visitor);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocumentVisitor, CommentsToText)
{
    s_instance->CommentsToText();
}

} // namespace gtest_test

void ExDocumentVisitor::FieldToText()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DocumentVisitor-compatible features.docx");
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExDocumentVisitor::FieldStructurePrinter>();
    
    // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
    // and then traverses all the node's children in a depth-first manner.
    // The visitor can read and modify each visited node.
    doc->Accept(visitor);
    
    std::cout << visitor->GetText() << std::endl;
    TestFieldToText(visitor);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocumentVisitor, FieldToText)
{
    s_instance->FieldToText();
}

} // namespace gtest_test

void ExDocumentVisitor::HeaderFooterToText()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DocumentVisitor-compatible features.docx");
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExDocumentVisitor::HeaderFooterStructurePrinter>();
    
    // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
    // and then traverses all the node's children in a depth-first manner.
    // The visitor can read and modify each visited node.
    doc->Accept(visitor);
    
    std::cout << visitor->GetText() << std::endl;
    
    // An alternative way of accessing a document's header/footers section-by-section is by accessing the collection.
    System::ArrayPtr<System::SharedPtr<Aspose::Words::HeaderFooter>> headerFooters = doc->get_FirstSection()->get_HeadersFooters()->ToArray();
    ASSERT_EQ(3, headerFooters->get_Length());
    TestHeaderFooterToText(visitor);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocumentVisitor, HeaderFooterToText)
{
    s_instance->HeaderFooterToText();
}

} // namespace gtest_test

void ExDocumentVisitor::EditableRangeToText()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DocumentVisitor-compatible features.docx");
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExDocumentVisitor::EditableRangeStructurePrinter>();
    
    // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
    // and then traverses all the node's children in a depth-first manner.
    // The visitor can read and modify each visited node.
    doc->Accept(visitor);
    
    std::cout << visitor->GetText() << std::endl;
    TestEditableRangeToText(visitor);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocumentVisitor, EditableRangeToText)
{
    s_instance->EditableRangeToText();
}

} // namespace gtest_test

void ExDocumentVisitor::FootnoteToText()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DocumentVisitor-compatible features.docx");
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExDocumentVisitor::FootnoteStructurePrinter>();
    
    // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
    // and then traverses all the node's children in a depth-first manner.
    // The visitor can read and modify each visited node.
    doc->Accept(visitor);
    
    std::cout << visitor->GetText() << std::endl;
    TestFootnoteToText(visitor);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocumentVisitor, FootnoteToText)
{
    s_instance->FootnoteToText();
}

} // namespace gtest_test

void ExDocumentVisitor::OfficeMathToText()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DocumentVisitor-compatible features.docx");
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExDocumentVisitor::OfficeMathStructurePrinter>();
    
    // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
    // and then traverses all the node's children in a depth-first manner.
    // The visitor can read and modify each visited node.
    doc->Accept(visitor);
    
    std::cout << visitor->GetText() << std::endl;
    TestOfficeMathToText(visitor);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocumentVisitor, OfficeMathToText)
{
    s_instance->OfficeMathToText();
}

} // namespace gtest_test

void ExDocumentVisitor::SmartTagToText()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Smart tags.doc");
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExDocumentVisitor::SmartTagStructurePrinter>();
    
    // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
    // and then traverses all the node's children in a depth-first manner.
    // The visitor can read and modify each visited node.
    doc->Accept(visitor);
    
    std::cout << visitor->GetText() << std::endl;
    TestSmartTagToText(visitor);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocumentVisitor, SmartTagToText)
{
    s_instance->SmartTagToText();
}

} // namespace gtest_test

void ExDocumentVisitor::StructuredDocumentTagToText()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DocumentVisitor-compatible features.docx");
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExDocumentVisitor::StructuredDocumentTagNodePrinter>();
    
    // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
    // and then traverses all the node's children in a depth-first manner.
    // The visitor can read and modify each visited node.
    doc->Accept(visitor);
    
    std::cout << visitor->GetText() << std::endl;
    TestStructuredDocumentTagToText(visitor);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocumentVisitor, StructuredDocumentTagToText)
{
    s_instance->StructuredDocumentTagToText();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
