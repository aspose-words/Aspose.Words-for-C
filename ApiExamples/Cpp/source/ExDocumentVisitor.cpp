// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentVisitor.h"

#include <testing/test_predicates.h>
#include <system/text/string_builder.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/date_time.h>
#include <system/console.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeStart.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Math/OfficeMath.h>
#include <Aspose.Words.Cpp/Model/Math/MathObjectType.h>
#include <Aspose.Words.Cpp/Model/Markup/SmartTag.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeStart.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeEnd.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRange.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/SubDocument.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

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

RTTI_INFO_IMPL_HASH(3594500859u, ::ApiExamples::ExDocumentVisitor::StructuredDocumentTagInfoPrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::StructuredDocumentTagInfoPrinter::StructuredDocumentTagInfoPrinter()
     : mVisitorIsInsideStructuredDocumentTag(false), mDocTraversalDepth(0)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideStructuredDocumentTag = false;
}

String ExDocumentVisitor::StructuredDocumentTagInfoPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::StructuredDocumentTagInfoPrinter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideStructuredDocumentTag)
    {
        IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
    }

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::StructuredDocumentTagInfoPrinter::VisitStructuredDocumentTagStart(SharedPtr<Aspose::Words::Markup::StructuredDocumentTag> sdt)
{
    IndentAndAppendLine(String(u"[StructuredDocumentTag start] Title: ") + sdt->get_Title());
    mDocTraversalDepth++;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::StructuredDocumentTagInfoPrinter::VisitStructuredDocumentTagEnd(SharedPtr<Aspose::Words::Markup::StructuredDocumentTag> sdt)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[StructuredDocumentTag end]");

    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::StructuredDocumentTagInfoPrinter::IndentAndAppendLine(String text)
{
    for (int i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }

    mBuilder->AppendLine(text);
}

System::Object::shared_members_type ApiExamples::ExDocumentVisitor::StructuredDocumentTagInfoPrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExDocumentVisitor::StructuredDocumentTagInfoPrinter::mBuilder", this->mBuilder);

    return result;
}

RTTI_INFO_IMPL_HASH(3245804126u, ::ApiExamples::ExDocumentVisitor::SmartTagInfoPrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::SmartTagInfoPrinter::SmartTagInfoPrinter() : mVisitorIsInsideSmartTag(false), mDocTraversalDepth(0)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideSmartTag = false;
}

String ExDocumentVisitor::SmartTagInfoPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::SmartTagInfoPrinter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideSmartTag)
    {
        IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
    }

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::SmartTagInfoPrinter::VisitSmartTagStart(SharedPtr<Aspose::Words::Markup::SmartTag> smartTag)
{
    IndentAndAppendLine(String(u"[SmartTag start] Name: ") + smartTag->get_Element());
    mDocTraversalDepth++;
    mVisitorIsInsideSmartTag = true;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::SmartTagInfoPrinter::VisitSmartTagEnd(SharedPtr<Aspose::Words::Markup::SmartTag> smartTag)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[SmartTag end]");
    mVisitorIsInsideSmartTag = false;

    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::SmartTagInfoPrinter::IndentAndAppendLine(String text)
{
    for (int i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }

    mBuilder->AppendLine(text);
}

System::Object::shared_members_type ApiExamples::ExDocumentVisitor::SmartTagInfoPrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExDocumentVisitor::SmartTagInfoPrinter::mBuilder", this->mBuilder);

    return result;
}

RTTI_INFO_IMPL_HASH(937070475u, ::ApiExamples::ExDocumentVisitor::OfficeMathInfoPrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::OfficeMathInfoPrinter::OfficeMathInfoPrinter() : mVisitorIsInsideOfficeMath(false)
    , mDocTraversalDepth(0)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideOfficeMath = false;
}

String ExDocumentVisitor::OfficeMathInfoPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::OfficeMathInfoPrinter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideOfficeMath)
    {
        IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
    }

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::OfficeMathInfoPrinter::VisitOfficeMathStart(SharedPtr<Aspose::Words::Math::OfficeMath> officeMath)
{
    IndentAndAppendLine(String(u"[OfficeMath start] Math object type: ") + System::ObjectExt::ToString(officeMath->get_MathObjectType()));
    mDocTraversalDepth++;
    mVisitorIsInsideOfficeMath = true;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::OfficeMathInfoPrinter::VisitOfficeMathEnd(SharedPtr<Aspose::Words::Math::OfficeMath> officeMath)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[OfficeMath end]");
    mVisitorIsInsideOfficeMath = false;

    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::OfficeMathInfoPrinter::IndentAndAppendLine(String text)
{
    for (int i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }

    mBuilder->AppendLine(text);
}

System::Object::shared_members_type ApiExamples::ExDocumentVisitor::OfficeMathInfoPrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExDocumentVisitor::OfficeMathInfoPrinter::mBuilder", this->mBuilder);

    return result;
}

RTTI_INFO_IMPL_HASH(3380632335u, ::ApiExamples::ExDocumentVisitor::FootnoteInfoPrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::FootnoteInfoPrinter::FootnoteInfoPrinter() : mVisitorIsInsideFootnote(false), mDocTraversalDepth(0)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideFootnote = false;
}

String ExDocumentVisitor::FootnoteInfoPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::FootnoteInfoPrinter::VisitFootnoteStart(SharedPtr<Aspose::Words::Footnote> footnote)
{
    IndentAndAppendLine(String(u"[Footnote start] Type: ") + System::ObjectExt::ToString(footnote->get_FootnoteType()));
    mDocTraversalDepth++;
    mVisitorIsInsideFootnote = true;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::FootnoteInfoPrinter::VisitFootnoteEnd(SharedPtr<Aspose::Words::Footnote> footnote)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Footnote end]");
    mVisitorIsInsideFootnote = false;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::FootnoteInfoPrinter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideFootnote)
    {
        IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
    }

    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::FootnoteInfoPrinter::IndentAndAppendLine(String text)
{
    for (int i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }

    mBuilder->AppendLine(text);
}

System::Object::shared_members_type ApiExamples::ExDocumentVisitor::FootnoteInfoPrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExDocumentVisitor::FootnoteInfoPrinter::mBuilder", this->mBuilder);

    return result;
}

RTTI_INFO_IMPL_HASH(3126765040u, ::ApiExamples::ExDocumentVisitor::EditableRangeInfoPrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::EditableRangeInfoPrinter::EditableRangeInfoPrinter() : mVisitorIsInsideEditableRange(false)
    , mDocTraversalDepth(0)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideEditableRange = false;
}

String ExDocumentVisitor::EditableRangeInfoPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::EditableRangeInfoPrinter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    // We want to print the contents of runs, but only if they are inside shapes, as they would be in the case of text boxes
    if (mVisitorIsInsideEditableRange)
    {
        IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
    }

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::EditableRangeInfoPrinter::VisitEditableRangeStart(SharedPtr<Aspose::Words::EditableRangeStart> editableRangeStart)
{
    IndentAndAppendLine(String(u"[EditableRange start] ID: ") + editableRangeStart->get_Id() + u" Owner: " + editableRangeStart->get_EditableRange()->get_SingleUser());
    mDocTraversalDepth++;
    mVisitorIsInsideEditableRange = true;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::EditableRangeInfoPrinter::VisitEditableRangeEnd(SharedPtr<Aspose::Words::EditableRangeEnd> editableRangeEnd)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[EditableRange end]");
    mVisitorIsInsideEditableRange = false;

    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::EditableRangeInfoPrinter::IndentAndAppendLine(String text)
{
    for (int i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }

    mBuilder->AppendLine(text);
}

System::Object::shared_members_type ApiExamples::ExDocumentVisitor::EditableRangeInfoPrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExDocumentVisitor::EditableRangeInfoPrinter::mBuilder", this->mBuilder);

    return result;
}

RTTI_INFO_IMPL_HASH(1356604423u, ::ApiExamples::ExDocumentVisitor::HeaderFooterInfoPrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::HeaderFooterInfoPrinter::HeaderFooterInfoPrinter() : mVisitorIsInsideHeaderFooter(false)
    , mDocTraversalDepth(0)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideHeaderFooter = false;
}

String ExDocumentVisitor::HeaderFooterInfoPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::HeaderFooterInfoPrinter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideHeaderFooter)
    {
        IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
    }

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::HeaderFooterInfoPrinter::VisitHeaderFooterStart(SharedPtr<Aspose::Words::HeaderFooter> headerFooter)
{
    IndentAndAppendLine(String(u"[HeaderFooter start] HeaderFooterType: ") + System::ObjectExt::ToString(headerFooter->get_HeaderFooterType()));
    mDocTraversalDepth++;
    mVisitorIsInsideHeaderFooter = true;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::HeaderFooterInfoPrinter::VisitHeaderFooterEnd(SharedPtr<Aspose::Words::HeaderFooter> headerFooter)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[HeaderFooter end]");
    mVisitorIsInsideHeaderFooter = false;

    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::HeaderFooterInfoPrinter::IndentAndAppendLine(String text)
{
    for (int i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }

    mBuilder->AppendLine(text);
}

System::Object::shared_members_type ApiExamples::ExDocumentVisitor::HeaderFooterInfoPrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExDocumentVisitor::HeaderFooterInfoPrinter::mBuilder", this->mBuilder);

    return result;
}

RTTI_INFO_IMPL_HASH(3636855599u, ::ApiExamples::ExDocumentVisitor::FieldInfoPrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::FieldInfoPrinter::FieldInfoPrinter() : mVisitorIsInsideField(false), mDocTraversalDepth(0)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideField = false;
}

String ExDocumentVisitor::FieldInfoPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::FieldInfoPrinter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideField)
    {
        IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
    }

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::FieldInfoPrinter::VisitFieldStart(SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart)
{
    IndentAndAppendLine(String(u"[Field start] FieldType: ") + System::ObjectExt::ToString(fieldStart->get_FieldType()));
    mDocTraversalDepth++;
    mVisitorIsInsideField = true;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::FieldInfoPrinter::VisitFieldEnd(SharedPtr<Aspose::Words::Fields::FieldEnd> fieldEnd)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Field end]");
    mVisitorIsInsideField = false;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::FieldInfoPrinter::VisitFieldSeparator(SharedPtr<Aspose::Words::Fields::FieldSeparator> fieldSeparator)
{
    IndentAndAppendLine(u"[FieldSeparator]");

    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::FieldInfoPrinter::IndentAndAppendLine(String text)
{
    for (int i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }

    mBuilder->AppendLine(text);
}

System::Object::shared_members_type ApiExamples::ExDocumentVisitor::FieldInfoPrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExDocumentVisitor::FieldInfoPrinter::mBuilder", this->mBuilder);

    return result;
}

RTTI_INFO_IMPL_HASH(1929097130u, ::ApiExamples::ExDocumentVisitor::CommentInfoPrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::CommentInfoPrinter::CommentInfoPrinter() : mVisitorIsInsideComment(false), mDocTraversalDepth(0)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideComment = false;
}

String ExDocumentVisitor::CommentInfoPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::CommentInfoPrinter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    if (mVisitorIsInsideComment)
    {
        IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
    }

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::CommentInfoPrinter::VisitCommentRangeStart(SharedPtr<Aspose::Words::CommentRangeStart> commentRangeStart)
{
    IndentAndAppendLine(String(u"[Comment range start] ID: ") + commentRangeStart->get_Id());
    mDocTraversalDepth++;
    mVisitorIsInsideComment = true;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::CommentInfoPrinter::VisitCommentRangeEnd(SharedPtr<Aspose::Words::CommentRangeEnd> commentRangeEnd)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Comment range end]");
    mVisitorIsInsideComment = false;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::CommentInfoPrinter::VisitCommentStart(SharedPtr<Aspose::Words::Comment> comment)
{
    IndentAndAppendLine(String::Format(u"[Comment start] For comment range ID {0}, By {1} on {2}",comment->get_Id(),comment->get_Author(),comment->get_DateTime()));
    mDocTraversalDepth++;
    mVisitorIsInsideComment = true;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::CommentInfoPrinter::VisitCommentEnd(SharedPtr<Aspose::Words::Comment> comment)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Comment end]");
    mVisitorIsInsideComment = false;

    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::CommentInfoPrinter::IndentAndAppendLine(String text)
{
    for (int i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }

    mBuilder->AppendLine(text);
}

System::Object::shared_members_type ApiExamples::ExDocumentVisitor::CommentInfoPrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExDocumentVisitor::CommentInfoPrinter::mBuilder", this->mBuilder);

    return result;
}

RTTI_INFO_IMPL_HASH(3340901531u, ::ApiExamples::ExDocumentVisitor::TableInfoPrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::TableInfoPrinter::TableInfoPrinter() : mVisitorIsInsideTable(false), mDocTraversalDepth(0)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
    mVisitorIsInsideTable = false;
}

String ExDocumentVisitor::TableInfoPrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableInfoPrinter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    // We want to print the contents of runs, but only if they consist of text from cells
    // So we are only interested in runs that are children of table nodes
    if (mVisitorIsInsideTable)
    {
        IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");
    }

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableInfoPrinter::VisitTableStart(SharedPtr<Aspose::Words::Tables::Table> table)
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

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableInfoPrinter::VisitTableEnd(SharedPtr<Aspose::Words::Tables::Table> table)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Table end]");
    mVisitorIsInsideTable = false;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableInfoPrinter::VisitRowStart(SharedPtr<Aspose::Words::Tables::Row> row)
{
    String rowContents = row->GetText().TrimEnd(MakeArray<char16_t>({u'\x0007', u' '})).Replace(u"\u0007", u", ");
    int rowWidth = row->IndexOf(row->get_LastCell()) + 1;
    int rowIndex = row->get_ParentTable()->IndexOf(row);
    String rowStatusInTable = row->get_IsFirstRow() && row->get_IsLastRow() ? u"only" : row->get_IsFirstRow() ? u"first" : row->get_IsLastRow() ? String(u"last") : String(u"");
    if (rowStatusInTable != u"")
    {
        rowStatusInTable = String::Format(u", the {0} row in this table,",rowStatusInTable);
    }

    IndentAndAppendLine(String::Format(u"[Row start] Row #{0}{1} width {2}, \"{3}\"",++rowIndex,rowStatusInTable,rowWidth,rowContents));
    mDocTraversalDepth++;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableInfoPrinter::VisitRowEnd(SharedPtr<Aspose::Words::Tables::Row> row)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Row end]");

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableInfoPrinter::VisitCellStart(SharedPtr<Aspose::Words::Tables::Cell> cell)
{
    SharedPtr<Row> row = cell->get_ParentRow();
    SharedPtr<Table> table = row->get_ParentTable();
    String cellStatusInRow = cell->get_IsFirstCell() && cell->get_IsLastCell() ? u"only" : cell->get_IsFirstCell() ? u"first" : cell->get_IsLastCell() ? String(u"last") : String(u"");
    if (cellStatusInRow != u"")
    {
        cellStatusInRow = String::Format(u", the {0} cell in this row",cellStatusInRow);
    }

    IndentAndAppendLine(String::Format(u"[Cell start] Row {0}, Col {1}{2}",table->IndexOf(row) + 1,row->IndexOf(cell) + 1,cellStatusInRow));
    mDocTraversalDepth++;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::TableInfoPrinter::VisitCellEnd(SharedPtr<Aspose::Words::Tables::Cell> cell)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Cell end]");
    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::TableInfoPrinter::IndentAndAppendLine(String text)
{
    for (int i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }

    mBuilder->AppendLine(text);
}

System::Object::shared_members_type ApiExamples::ExDocumentVisitor::TableInfoPrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExDocumentVisitor::TableInfoPrinter::mBuilder", this->mBuilder);

    return result;
}

RTTI_INFO_IMPL_HASH(4082220962u, ::ApiExamples::ExDocumentVisitor::DocStructurePrinter, ThisTypeBaseTypesInfo);

ExDocumentVisitor::DocStructurePrinter::DocStructurePrinter() : mDocTraversalDepth(0)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
}

String ExDocumentVisitor::DocStructurePrinter::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitDocumentStart(SharedPtr<Aspose::Words::Document> doc)
{
    int childNodeCount = doc->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count();

    // A Document node is at the root of every document, so if we let a document accept a visitor, this will be the first visitor action to be carried out
    IndentAndAppendLine(String(u"[Document start] Child nodes: ") + childNodeCount);
    mDocTraversalDepth++;

    // Let the visitor continue visiting other nodes
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitDocumentEnd(SharedPtr<Aspose::Words::Document> doc)
{
    // If we let a document accept a visitor, this will be the last visitor action to be carried out
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Document end]");

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitSectionStart(SharedPtr<Aspose::Words::Section> section)
{
    // Get the index of our section within the document
    SharedPtr<NodeCollection> docSections = section->get_Document()->GetChildNodes(Aspose::Words::NodeType::Section, false);
    int sectionIndex = docSections->IndexOf(section);

    IndentAndAppendLine(String(u"[Section start] Section index: ") + sectionIndex);
    mDocTraversalDepth++;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitSectionEnd(SharedPtr<Aspose::Words::Section> section)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Section end]");

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitBodyStart(SharedPtr<Aspose::Words::Body> body)
{
    int paragraphCount = body->get_Paragraphs()->get_Count();
    IndentAndAppendLine(String(u"[Body start] Paragraphs: ") + paragraphCount);
    mDocTraversalDepth++;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitBodyEnd(SharedPtr<Aspose::Words::Body> body)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Body end]");

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitParagraphStart(SharedPtr<Aspose::Words::Paragraph> paragraph)
{
    IndentAndAppendLine(u"[Paragraph start]");
    mDocTraversalDepth++;

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitParagraphEnd(SharedPtr<Aspose::Words::Paragraph> paragraph)
{
    mDocTraversalDepth--;
    IndentAndAppendLine(u"[Paragraph end]");

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitRun(SharedPtr<Aspose::Words::Run> run)
{
    IndentAndAppendLine(String(u"[Run] \"") + run->GetText() + u"\"");

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExDocumentVisitor::DocStructurePrinter::VisitSubDocument(SharedPtr<Aspose::Words::SubDocument> subDocument)
{
    IndentAndAppendLine(u"[SubDocument]");

    return Aspose::Words::VisitorAction::Continue;
}

void ExDocumentVisitor::DocStructurePrinter::IndentAndAppendLine(String text)
{
    for (int i = 0; i < mDocTraversalDepth; i++)
    {
        mBuilder->Append(u"|  ");
    }

    mBuilder->AppendLine(text);
}

System::Object::shared_members_type ApiExamples::ExDocumentVisitor::DocStructurePrinter::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExDocumentVisitor::DocStructurePrinter::mBuilder", this->mBuilder);

    return result;
}

void ExDocumentVisitor::TestDocStructureToText(SharedPtr<ExDocumentVisitor::DocStructurePrinter> visitor)
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

void ExDocumentVisitor::TestTableToText(SharedPtr<ExDocumentVisitor::TableInfoPrinter> visitor)
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

void ExDocumentVisitor::TestCommentsToText(SharedPtr<ExDocumentVisitor::CommentInfoPrinter> visitor)
{
    String visitorText = visitor->GetText();

    ASSERT_TRUE(visitorText.Contains(u"[Comment range start]"));
    ASSERT_TRUE(visitorText.Contains(u"[Comment range end]"));
    ASSERT_TRUE(visitorText.Contains(u"[Comment start]"));
    ASSERT_TRUE(visitorText.Contains(u"[Comment end]"));
    ASSERT_TRUE(visitorText.Contains(u"[Run]"));
}

void ExDocumentVisitor::TestFieldToText(SharedPtr<ExDocumentVisitor::FieldInfoPrinter> visitor)
{
    String visitorText = visitor->GetText();

    ASSERT_TRUE(visitorText.Contains(u"[Field start]"));
    ASSERT_TRUE(visitorText.Contains(u"[Field end]"));
    ASSERT_TRUE(visitorText.Contains(u"[FieldSeparator]"));
    ASSERT_TRUE(visitorText.Contains(u"[Run]"));
}

void ExDocumentVisitor::TestHeaderFooterToText(SharedPtr<ExDocumentVisitor::HeaderFooterInfoPrinter> visitor)
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

void ExDocumentVisitor::TestEditableRangeToText(SharedPtr<ExDocumentVisitor::EditableRangeInfoPrinter> visitor)
{
    String visitorText = visitor->GetText();

    ASSERT_TRUE(visitorText.Contains(u"[EditableRange start]"));
    ASSERT_TRUE(visitorText.Contains(u"[EditableRange end]"));
    ASSERT_TRUE(visitorText.Contains(u"[Run]"));
}

void ExDocumentVisitor::TestFootnoteToText(SharedPtr<ExDocumentVisitor::FootnoteInfoPrinter> visitor)
{
    String visitorText = visitor->GetText();

    ASSERT_TRUE(visitorText.Contains(u"[Footnote start] Type: Footnote"));
    ASSERT_TRUE(visitorText.Contains(u"[Footnote end]"));
    ASSERT_TRUE(visitorText.Contains(u"[Run]"));
}

void ExDocumentVisitor::TestOfficeMathToText(SharedPtr<ExDocumentVisitor::OfficeMathInfoPrinter> visitor)
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

void ExDocumentVisitor::TestSmartTagToText(SharedPtr<ExDocumentVisitor::SmartTagInfoPrinter> visitor)
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

void ExDocumentVisitor::TestStructuredDocumentTagToText(SharedPtr<ExDocumentVisitor::StructuredDocumentTagInfoPrinter> visitor)
{
    String visitorText = visitor->GetText();

    ASSERT_TRUE(visitorText.Contains(u"[StructuredDocumentTag start]"));
    ASSERT_TRUE(visitorText.Contains(u"[StructuredDocumentTag end]"));
}

namespace gtest_test
{

class ExDocumentVisitor : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExDocumentVisitor> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExDocumentVisitor>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExDocumentVisitor> ExDocumentVisitor::s_instance;

} // namespace gtest_test

void ExDocumentVisitor::DocStructureToText()
{
    // Open the document that has nodes we want to print the info of
    auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");

    // Create an object that inherits from the DocumentVisitor class
    auto visitor = MakeObject<ExDocumentVisitor::DocStructurePrinter>();

    // Accepting a visitor lets it start traversing the nodes in the document,
    // starting with the node that accepted it to then recursively visit every child
    doc->Accept(visitor);

    // Once the visiting is complete, we can retrieve the result of the operation,
    // that in this example, has accumulated in the visitor
    System::Console::WriteLine(visitor->GetText());
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
    // Open the document that has tables we want to print the info of
    auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");

    // Create an object that inherits from the DocumentVisitor class
    auto visitor = MakeObject<ExDocumentVisitor::TableInfoPrinter>();

    // Accepting a visitor lets it start traversing the nodes in the document,
    // starting with the node that accepted it to then recursively visit every child
    doc->Accept(visitor);

    // Once the visiting is complete, we can retrieve the result of the operation,
    // that in this example, has accumulated in the visitor
    System::Console::WriteLine(visitor->GetText());
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
    // Open the document that has comments/comment ranges we want to print the info of
    auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");

    // Create an object that inherits from the DocumentVisitor class
    auto visitor = MakeObject<ExDocumentVisitor::CommentInfoPrinter>();

    // Accepting a visitor lets it start traversing the nodes in the document,
    // starting with the node that accepted it to then recursively visit every child
    doc->Accept(visitor);

    // Once the visiting is complete, we can retrieve the result of the operation,
    // that in this example, has accumulated in the visitor
    System::Console::WriteLine(visitor->GetText());
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
    // Open the document that has fields that we want to print the info of
    auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");

    // Create an object that inherits from the DocumentVisitor class
    auto visitor = MakeObject<ExDocumentVisitor::FieldInfoPrinter>();

    // Accepting a visitor lets it start traversing the nodes in the document,
    // starting with the node that accepted it to then recursively visit every child
    doc->Accept(visitor);

    // Once the visiting is complete, we can retrieve the result of the operation,
    // that in this example, has accumulated in the visitor
    System::Console::WriteLine(visitor->GetText());
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
    // Open the document that has headers and/or footers we want to print the info of
    auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");

    // Create an object that inherits from the DocumentVisitor class
    auto visitor = MakeObject<ExDocumentVisitor::HeaderFooterInfoPrinter>();

    // Accepting a visitor lets it start traversing the nodes in the document,
    // starting with the node that accepted it to then recursively visit every child
    doc->Accept(visitor);

    // Once the visiting is complete, we can retrieve the result of the operation,
    // that in this example, has accumulated in the visitor
    System::Console::WriteLine(visitor->GetText());

    // An alternative way of visiting a document's header/footers section-by-section is by accessing the collection
    // We can also turn it into an array
    ArrayPtr<SharedPtr<HeaderFooter>> headerFooters = doc->get_FirstSection()->get_HeadersFooters()->ToArray();
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
    // Open the document that has editable ranges we want to print the info of
    auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");

    // Create an object that inherits from the DocumentVisitor class
    auto visitor = MakeObject<ExDocumentVisitor::EditableRangeInfoPrinter>();

    // Accepting a visitor lets it start traversing the nodes in the document,
    // starting with the node that accepted it to then recursively visit every child
    doc->Accept(visitor);

    auto p = MakeObject<Paragraph>(doc);
    p->AppendChild(MakeObject<Run>(doc, u"Paragraph with editable range text."));

    // Once the visiting is complete, we can retrieve the result of the operation,
    // that in this example, has accumulated in the visitor
    System::Console::WriteLine(visitor->GetText());
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
    // Open the document that has footnotes we want to print the info of
    auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");

    // Create an object that inherits from the DocumentVisitor class
    auto visitor = MakeObject<ExDocumentVisitor::FootnoteInfoPrinter>();

    // Accepting a visitor lets it start traversing the nodes in the document,
    // starting with the node that accepted it to then recursively visit every child
    doc->Accept(visitor);

    // Once the visiting is complete, we can retrieve the result of the operation,
    // that in this example, has accumulated in the visitor
    System::Console::WriteLine(visitor->GetText());
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
    // Open the document that has office math objects we want to print the info of
    auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");

    // Create an object that inherits from the DocumentVisitor class
    auto visitor = MakeObject<ExDocumentVisitor::OfficeMathInfoPrinter>();

    // Accepting a visitor lets it start traversing the nodes in the document,
    // starting with the node that accepted it to then recursively visit every child
    doc->Accept(visitor);

    // Once the visiting is complete, we can retrieve the result of the operation,
    // that in this example, has accumulated in the visitor
    System::Console::WriteLine(visitor->GetText());
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
    // Open the document that has smart tags we want to print the info of
    auto doc = MakeObject<Document>(MyDir + u"Smart tags.doc");

    // Create an object that inherits from the DocumentVisitor class
    auto visitor = MakeObject<ExDocumentVisitor::SmartTagInfoPrinter>();

    // Accepting a visitor lets it start traversing the nodes in the document,
    // starting with the node that accepted it to then recursively visit every child
    doc->Accept(visitor);

    // Once the visiting is complete, we can retrieve the result of the operation,
    // that in this example, has accumulated in the visitor
    System::Console::WriteLine(visitor->GetText());
    TestSmartTagToText(visitor);
    //ExEnd
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
    // Open the document that has structured document tags we want to print the info of
    auto doc = MakeObject<Document>(MyDir + u"DocumentVisitor-compatible features.docx");

    // Create an object that inherits from the DocumentVisitor class
    auto visitor = MakeObject<ExDocumentVisitor::StructuredDocumentTagInfoPrinter>();

    // Accepting a visitor lets it start traversing the nodes in the document,
    // starting with the node that accepted it to then recursively visit every child
    doc->Accept(visitor);

    // Once the visiting is complete, we can retrieve the result of the operation,
    // that in this example, has accumulated in the visitor
    System::Console::WriteLine(visitor->GetText());
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
