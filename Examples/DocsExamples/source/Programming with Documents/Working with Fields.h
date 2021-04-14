#pragma once

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldUpdateCultureSource.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldIf.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldIfComparisonResult.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldAuthor.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/EquationsAndFormulas/FieldAdvance.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldTA.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldToa.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldAsk.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldIncludeText.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldAddressBlock.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldMergeField.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Other/FieldUnknown.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldUpdateCultureProvider.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MappedDataFieldCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <system/action.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/collections/list.h>
#include <system/date_time.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/globalization/culture_info.h>
#include <system/globalization/date_time_format_info.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/scope_guard.h>
#include <system/text/regularexpressions/group.h>
#include <system/text/regularexpressions/group_collection.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <system/threading/thread.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithFields : public DocsExamplesBase
{
public:
    /// <summary>
    /// Represents a facade object for a merge field in a Microsoft Word document.
    /// </summary>
    class MergeField : public System::Object
    {
    public:
        /// <summary>
        /// Gets the name of the merge field.
        /// </summary>
        String get_Name()
        {
            return (System::DynamicCast<FieldStart>(mFieldStart))->GetField()->get_Result().Replace(u"«", u"").Replace(u"»", u"");
        }

        /// <summary>
        /// Sets the name of the merge field.
        /// </summary>
        void set_Name(String value)
        {
            // Merge field name is stored in the field result which is a Run
            // node between field separator and field end.
            auto fieldResult = System::DynamicCast<Aspose::Words::Run>(mFieldSeparator->get_NextSibling());
            fieldResult->set_Text(String::Format(u"«{0}»", value));

            // But sometimes the field result can consist of more than one run, delete these runs.
            RemoveSameParent(fieldResult->get_NextSibling(), mFieldEnd);

            UpdateFieldCode(value);
        }

        MergeField(SharedPtr<FieldStart> fieldStart)
            : gRegex(MakeObject<System::Text::RegularExpressions::Regex>(u"\\s*(?<start>MERGEFIELD\\s|)(\\s|)(?<name>\\S+)\\s+"))
        {
            // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            System::Details::ThisProtector __local_self_ref(this);
            //---------------------------------------------------------Self Reference

            if (fieldStart == nullptr)
            {
                throw System::ArgumentNullException(u"fieldStart");
            }
            if (fieldStart->get_FieldType() != FieldType::FieldMergeField)
            {
                throw System::ArgumentException(u"Field start type must be FieldMergeField.");
            }

            mFieldStart = fieldStart;

            // Find the field separator node.
            mFieldSeparator = fieldStart->GetField()->get_Separator();
            if (mFieldSeparator == nullptr)
            {
                throw System::InvalidOperationException(u"Cannot find field separator.");
            }

            mFieldEnd = fieldStart->GetField()->get_End();
        }

    private:
        SharedPtr<Node> mFieldStart;
        SharedPtr<Node> mFieldSeparator;
        SharedPtr<Node> mFieldEnd;
        SharedPtr<System::Text::RegularExpressions::Regex> gRegex;

        void UpdateFieldCode(String fieldName)
        {
            // Field code is stored in a Run node between field start and field separator.
            auto fieldCode = System::DynamicCast<Aspose::Words::Run>(mFieldStart->get_NextSibling());

            SharedPtr<System::Text::RegularExpressions::Match> match =
                gRegex->Match((System::DynamicCast<FieldStart>(mFieldStart))->GetField()->GetFieldCode());

            String newFieldCode = String::Format(u" {0}{1} ", match->get_Groups()->idx_get(u"start")->get_Value(), fieldName);
            fieldCode->set_Text(newFieldCode);

            // But sometimes the field code can consist of more than one run, delete these runs.
            RemoveSameParent(fieldCode->get_NextSibling(), mFieldSeparator);
        }

        /// <summary>
        /// Removes nodes from start up to but not including the end node.
        /// Start and end are assumed to have the same parent.
        /// </summary>
        void RemoveSameParent(SharedPtr<Node> startNode, SharedPtr<Node> endNode)
        {
            if (endNode != nullptr && startNode->get_ParentNode() != endNode->get_ParentNode())
            {
                throw System::ArgumentException(u"Start and end nodes are expected to have the same parent.");
            }

            SharedPtr<Node> curChild = startNode;
            while (curChild != nullptr && curChild != endNode)
            {
                SharedPtr<Node> nextChild = curChild->get_NextSibling();
                curChild->Remove();
                curChild = nextChild;
            }
        }
    };

private:
    class FieldUpdateCultureProvider : public IFieldUpdateCultureProvider
    {
    public:
        SharedPtr<System::Globalization::CultureInfo> GetCulture(String name, SharedPtr<Field> field) override
        {
            ASPOSE_UNUSED(field);
            if (name == u"ru-RU")
            {
                auto culture = MakeObject<System::Globalization::CultureInfo>(name, false);
                SharedPtr<System::Globalization::DateTimeFormatInfo> format = culture->get_DateTimeFormat();
                format->set_MonthNames(MakeArray<String>({u"месяц 1", u"месяц 2", u"месяц 3", u"месяц 4", u"месяц 5", u"месяц 6", u"месяц 7", u"месяц 8",
                                                          u"месяц 9", u"месяц 10", u"месяц 11", u"месяц 12", u""}));
                format->set_MonthGenitiveNames(format->get_MonthNames());
                format->set_AbbreviatedMonthNames(MakeArray<String>(
                    {u"мес 1", u"мес 2", u"мес 3", u"мес 4", u"мес 5", u"мес 6", u"мес 7", u"мес 8", u"мес 9", u"мес 10", u"мес 11", u"мес 12", u""}));
                format->set_AbbreviatedMonthGenitiveNames(format->get_AbbreviatedMonthNames());
                format->set_DayNames(MakeArray<String>(
                    {u"день недели 7", u"день недели 1", u"день недели 2", u"день недели 3", u"день недели 4", u"день недели 5", u"день недели 6"}));
                format->set_AbbreviatedDayNames(MakeArray<String>({u"день 7", u"день 1", u"день 2", u"день 3", u"день 4", u"день 5", u"день 6"}));
                format->set_ShortestDayNames(MakeArray<String>({u"д7", u"д1", u"д2", u"д3", u"д4", u"д5", u"д6"}));
                format->set_AMDesignator(u"До полудня");
                format->set_PMDesignator(u"После полудня");
                const String pattern = u"yyyy MM (MMMM) dd (dddd) hh:mm:ss tt";
                format->set_LongDatePattern(pattern);
                format->set_LongTimePattern(pattern);
                format->set_ShortDatePattern(pattern);
                format->set_ShortTimePattern(pattern);
                return culture;
            }
            else if (name == u"en-US")
            {
                return MakeObject<System::Globalization::CultureInfo>(name, false);
            }
            else
            {
                return nullptr;
            }
        }
    };

public:
    void ChangeFieldUpdateCultureSource()
    {
        //ExStart:ChangeFieldUpdateCultureSource
        //ExStart:DocumentBuilderInsertField
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert content with German locale.
        builder->get_Font()->set_LocaleId(1031);
        builder->InsertField(u"MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
        builder->Write(u" - ");
        builder->InsertField(u"MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");
        //ExEnd:DocumentBuilderInsertField

        // Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from
        // set the culture used during field update to the culture used by the field.
        doc->get_FieldOptions()->set_FieldUpdateCultureSource(FieldUpdateCultureSource::FieldCode);
        doc->get_MailMerge()->Execute(MakeArray<String>({u"Date2"}),
                                      MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<System::DateTime>(System::DateTime(2011, 1, 1))}));

        doc->Save(ArtifactsDir + u"WorkingWithFields.ChangeFieldUpdateCultureSource.docx");
        //ExEnd:ChangeFieldUpdateCultureSource
    }

    void SpecifyLocaleAtFieldLevel()
    {
        //ExStart:SpecifylocaleAtFieldlevel
        auto builder = MakeObject<DocumentBuilder>();

        SharedPtr<Field> field = builder->InsertField(FieldType::FieldDate, true);
        field->set_LocaleId(1049);

        builder->get_Document()->Save(ArtifactsDir + u"WorkingWithFields.SpecifylocaleAtFieldlevel.docx");
        //ExEnd:SpecifylocaleAtFieldlevel
    }

    void ReplaceHyperlinks()
    {
        //ExStart:ReplaceHyperlinks
        auto doc = MakeObject<Document>(MyDir + u"Hyperlinks.docx");

        for (auto field : System::IterateOver(doc->get_Range()->get_Fields()))
        {
            if (field->get_Type() == FieldType::FieldHyperlink)
            {
                auto hyperlink = System::DynamicCast<FieldHyperlink>(field);

                // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                if (hyperlink->get_SubAddress() != nullptr)
                {
                    continue;
                }

                hyperlink->set_Address(u"http://www.aspose.com");
                hyperlink->set_Result(u"Aspose - The .NET & Java Component Publisher");
            }
        }

        doc->Save(ArtifactsDir + u"WorkingWithFields.ReplaceHyperlinks.docx");
        //ExEnd:ReplaceHyperlinks
    }

    void RenameMergeFields()
    {
        //ExStart:RenameMergeFields
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u"MERGEFIELD MyMergeField1 \\* MERGEFORMAT");
        builder->InsertField(u"MERGEFIELD MyMergeField2 \\* MERGEFORMAT");

        // Select all field start nodes so we can find the merge fields.
        SharedPtr<NodeCollection> fieldStarts = doc->GetChildNodes(NodeType::FieldStart, true);
        for (auto fieldStart : System::IterateOver<FieldStart>(fieldStarts))
        {
            if (fieldStart->get_FieldType() == FieldType::FieldMergeField)
            {
                auto mergeField = MakeObject<WorkingWithFields::MergeField>(fieldStart);
                mergeField->set_Name(mergeField->get_Name() + u"_Renamed");
            }
        }

        doc->Save(ArtifactsDir + u"WorkingWithFields.RenameMergeFields.doc");
        //ExEnd:RenameMergeFields
    }

    void RemoveField()
    {
        //ExStart:RemoveField
        auto doc = MakeObject<Document>(MyDir + u"Various fields.docx");

        SharedPtr<Field> field = doc->get_Range()->get_Fields()->idx_get(0);
        field->Remove();
        //ExEnd:RemoveField
    }

    void InsertTOAFieldWithoutDocumentBuilder()
    {
        //ExStart:InsertTOAFieldWithoutDocumentBuilder
        auto doc = MakeObject<Document>();
        auto para = MakeObject<Paragraph>(doc);

        // We want to insert TA and TOA fields like this:
        // { TA  \c 1 \l "Value 0" }
        // { TOA  \c 1 }

        auto fieldTA = System::DynamicCast<FieldTA>(para->AppendField(FieldType::FieldTOAEntry, false));
        fieldTA->set_EntryCategory(u"1");
        fieldTA->set_LongCitation(u"Value 0");

        doc->get_FirstSection()->get_Body()->AppendChild(para);

        para = MakeObject<Paragraph>(doc);

        auto fieldToa = System::DynamicCast<FieldToa>(para->AppendField(FieldType::FieldTOA, false));
        fieldToa->set_EntryCategory(u"1");
        doc->get_FirstSection()->get_Body()->AppendChild(para);

        fieldToa->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertTOAFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertTOAFieldWithoutDocumentBuilder
    }

    void InsertNestedFields()
    {
        //ExStart:InsertNestedFields
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        for (int i = 0; i < 5; i++)
        {
            builder->InsertBreak(BreakType::PageBreak);
        }

        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);

        // We want to insert a field like this:
        // { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
        SharedPtr<Field> field = builder->InsertField(u"IF ");
        builder->MoveTo(field->get_Separator());
        builder->InsertField(u"PAGE");
        builder->Write(u" <> ");
        builder->InsertField(u"NUMPAGES");
        builder->Write(u" \"See Next Page\" \"Last Page\" ");

        field->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertNestedFields.docx");
        //ExEnd:InsertNestedFields
    }

    void InsertMergeFieldUsingDOM()
    {
        //ExStart:InsertMergeFieldUsingDOM
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto para = System::DynamicCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(0));

        builder->MoveTo(para);

        // We want to insert a merge field like this:
        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }

        auto field = System::DynamicCast<FieldMergeField>(builder->InsertField(FieldType::FieldMergeField, false));

        // { " MERGEFIELD Test1" }
        field->set_FieldName(u"Test1");

        // { " MERGEFIELD Test1 \\b Test2" }
        field->set_TextBefore(u"Test2");

        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 }
        field->set_TextAfter(u"Test3");

        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m" }
        field->set_IsMapped(true);

        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
        field->set_IsVerticalFormatting(true);

        // Finally update this merge field
        field->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertMergeFieldUsingDOM.docx");
        //ExEnd:InsertMergeFieldUsingDOM
    }

    void InsertMailMergeAddressBlockFieldUsingDOM()
    {
        //ExStart:InsertMailMergeAddressBlockFieldUsingDOM
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto para = System::DynamicCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(0));

        builder->MoveTo(para);

        // We want to insert a mail merge address block like this:
        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }

        auto field = System::DynamicCast<FieldAddressBlock>(builder->InsertField(FieldType::FieldAddressBlock, false));

        // { ADDRESSBLOCK \\c 1" }
        field->set_IncludeCountryOrRegionName(u"1");

        // { ADDRESSBLOCK \\c 1 \\d" }
        field->set_FormatAddressOnCountryOrRegion(true);

        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 }
        field->set_ExcludedCountryOrRegionName(u"Test2");

        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 }
        field->set_NameAndAddressFormat(u"Test3");

        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }
        field->set_LanguageId(u"Test 4");

        field->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertMailMergeAddressBlockFieldUsingDOM.docx");
        //ExEnd:InsertMailMergeAddressBlockFieldUsingDOM
    }

    void InsertFieldIncludeTextWithoutDocumentBuilder()
    {
        //ExStart:InsertFieldIncludeTextWithoutDocumentBuilder
        auto doc = MakeObject<Document>();

        auto para = MakeObject<Paragraph>(doc);

        // We want to insert an INCLUDETEXT field like this:
        // { INCLUDETEXT  "file path" }

        auto fieldIncludeText = System::DynamicCast<FieldIncludeText>(para->AppendField(FieldType::FieldIncludeText, false));
        fieldIncludeText->set_BookmarkName(u"bookmark");
        fieldIncludeText->set_SourceFullName(MyDir + u"IncludeText.docx");

        doc->get_FirstSection()->get_Body()->AppendChild(para);

        fieldIncludeText->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertIncludeFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertFieldIncludeTextWithoutDocumentBuilder
    }

    void InsertFieldNone()
    {
        //ExStart:InsertFieldNone
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto field = System::DynamicCast<FieldUnknown>(builder->InsertField(FieldType::FieldNone, false));

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertFieldNone.docx");
        //ExEnd:InsertFieldNone
    }

    void InsertField()
    {
        //ExStart:InsertField
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u"MERGEFIELD MyFieldName \\* MERGEFORMAT");

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertField.docx");
        //ExEnd:InsertField
    }

    void InsertAuthorField()
    {
        //ExStart:InsertAuthorField
        auto doc = MakeObject<Document>();

        auto para = System::DynamicCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(0));

        // We want to insert an AUTHOR field like this:
        // { AUTHOR Test1 }

        auto field = System::DynamicCast<FieldAuthor>(para->AppendField(FieldType::FieldAuthor, false));
        field->set_AuthorName(u"Test1");
        // { AUTHOR Test1 }

        field->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertAuthorField.docx");
        //ExEnd:InsertAuthorField
    }

    void InsertASKFieldWithOutDocumentBuilder()
    {
        //ExStart:InsertASKFieldWithOutDocumentBuilder
        auto doc = MakeObject<Document>();

        auto para = System::DynamicCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(0));

        // We want to insert an Ask field like this:
        // { ASK \"Test 1\" Test2 \\d Test3 \\o }

        auto field = System::DynamicCast<FieldAsk>(para->AppendField(FieldType::FieldAsk, false));

        // { ASK \"Test 1\" " }
        field->set_BookmarkName(u"Test 1");

        // { ASK \"Test 1\" Test2 }
        field->set_PromptText(u"Test2");

        // { ASK \"Test 1\" Test2 \\d Test3 }
        field->set_DefaultResponse(u"Test3");

        // { ASK \"Test 1\" Test2 \\d Test3 \\o }
        field->set_PromptOnceOnMailMerge(true);

        field->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertASKFieldWithOutDocumentBuilder.docx");
        //ExEnd:InsertASKFieldWithOutDocumentBuilder
    }

    void InsertAdvanceFieldWithOutDocumentBuilder()
    {
        //ExStart:InsertAdvanceFieldWithOutDocumentBuilder
        auto doc = MakeObject<Document>();

        auto para = System::DynamicCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(0));

        // We want to insert an Advance field like this:
        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }

        auto field = System::DynamicCast<FieldAdvance>(para->AppendField(FieldType::FieldAdvance, false));

        // { ADVANCE \\d 10 " }
        field->set_DownOffset(u"10");

        // { ADVANCE \\d 10 \\l 10 }
        field->set_LeftOffset(u"10");

        // { ADVANCE \\d 10 \\l 10 \\r -3.3 }
        field->set_RightOffset(u"-3.3");

        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 }
        field->set_UpOffset(u"0");

        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 }
        field->set_HorizontalPosition(u"100");

        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }
        field->set_VerticalPosition(u"100");

        field->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertAdvanceFieldWithOutDocumentBuilder.docx");
        //ExEnd:InsertAdvanceFieldWithOutDocumentBuilder
    }

    void GetMailMergeFieldNames()
    {
        //ExStart:GetFieldNames
        auto doc = MakeObject<Document>();

        ArrayPtr<String> fieldNames = doc->get_MailMerge()->GetFieldNames();
        //ExEnd:GetFieldNames
        std::cout << (String(u"\nDocument have ") + fieldNames->get_Length() + u" fields.") << std::endl;
    }

    void MappedDataFields()
    {
        //ExStart:MappedDataFields
        auto doc = MakeObject<Document>();

        doc->get_MailMerge()->get_MappedDataFields()->Add(u"MyFieldName_InDocument", u"MyFieldName_InDataSource");
        //ExEnd:MappedDataFields
    }

    void DeleteFields()
    {
        //ExStart:DeleteFields
        auto doc = MakeObject<Document>();

        doc->get_MailMerge()->DeleteFields();
        //ExEnd:DeleteFields
    }

    void FieldUpdateCulture()
    {
        //ExStart:FieldUpdateCultureProvider
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(FieldType::FieldTime, true);

        doc->get_FieldOptions()->set_FieldUpdateCultureSource(FieldUpdateCultureSource::FieldCode);
        doc->get_FieldOptions()->set_FieldUpdateCultureProvider(MakeObject<WorkingWithFields::FieldUpdateCultureProvider>());

        doc->Save(ArtifactsDir + u"WorkingWithFields.FieldUpdateCulture.pdf");
        //ExEnd:FieldUpdateCultureProvider
    }

    void FieldDisplayResults()
    {
        //ExStart:FieldDisplayResults
        //ExStart:UpdateDocFields
        auto document = MakeObject<Document>(MyDir + u"Various fields.docx");

        document->UpdateFields();
        //ExEnd:UpdateDocFields

        for (auto field : System::IterateOver(document->get_Range()->get_Fields()))
        {
            std::cout << field->get_DisplayResult() << std::endl;
        }
        //ExEnd:FieldDisplayResults
    }

    void EvaluateIFCondition()
    {
        //ExStart:EvaluateIFCondition
        auto builder = MakeObject<DocumentBuilder>();

        auto field = System::DynamicCast<FieldIf>(builder->InsertField(u"IF 1 = 1", nullptr));
        FieldIfComparisonResult actualResult = field->EvaluateCondition();

        std::cout << System::EnumGetName(actualResult) << std::endl;
        //ExEnd:EvaluateIFCondition
    }

    void ConvertFieldsInParagraph()
    {
        //ExStart:ConvertFieldsInParagraph
        auto doc = MakeObject<Document>(MyDir + u"Linked fields.docx");

        // Pass the appropriate parameters to convert all IF fields to text that are encountered only in the last
        // paragraph of the document.
        doc->get_FirstSection()
            ->get_Body()
            ->get_LastParagraph()
            ->get_Range()
            ->get_Fields()
            ->LINQ_Where([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldIf; })
            ->LINQ_ToList()
            ->ForEach(std::function<void(SharedPtr<Field>)>([](SharedPtr<Field> f) { f->Unlink(); }));

        doc->Save(ArtifactsDir + u"WorkingWithFields.TestFile.docx");
        //ExEnd:ConvertFieldsInParagraph
    }

    void ConvertFieldsInDocument()
    {
        //ExStart:ConvertFieldsInDocument
        auto doc = MakeObject<Document>(MyDir + u"Linked fields.docx");

        // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to text.
        doc->get_Range()
            ->get_Fields()
            ->LINQ_Where([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldIf; })
            ->LINQ_ToList()
            ->ForEach(std::function<void(SharedPtr<Field> f)>([](SharedPtr<Field> f) { f->Unlink(); }));

        // Save the document with fields transformed to disk
        doc->Save(ArtifactsDir + u"WorkingWithFields.ConvertFieldsInDocument.docx");
        //ExEnd:ConvertFieldsInDocument
    }

    void ConvertFieldsInBody()
    {
        //ExStart:ConvertFieldsInBody
        auto doc = MakeObject<Document>(MyDir + u"Linked fields.docx");

        // Pass the appropriate parameters to convert PAGE fields encountered to text only in the body of the first section.
        doc->get_FirstSection()
            ->get_Body()
            ->get_Range()
            ->get_Fields()
            ->LINQ_Where([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldPage; })
            ->LINQ_ToList()
            ->ForEach(std::function<void(SharedPtr<Field> f)>([](SharedPtr<Field> f) { f->Unlink(); }));

        doc->Save(ArtifactsDir + u"WorkingWithFields.ConvertFieldsInBody.docx");
        //ExEnd:ConvertFieldsInBody
    }

    void ChangeLocale()
    {
        //ExStart:ChangeLocale
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u"MERGEFIELD Date");

        // Store the current culture so it can be set back once mail merge is complete.
        SharedPtr<System::Globalization::CultureInfo> currentCulture = System::Threading::Thread::get_CurrentThread()->get_CurrentCulture();
        // Set to German language so dates and numbers are formatted using this culture during mail merge.
        System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(MakeObject<System::Globalization::CultureInfo>(u"de-DE"));

        doc->get_MailMerge()->Execute(MakeArray<String>({u"Date"}),
                                      MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<System::DateTime>(System::DateTime::get_Now())}));

        System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(currentCulture);

        doc->Save(ArtifactsDir + u"WorkingWithFields.ChangeLocale.docx");
        //ExEnd:ChangeLocale
    }
};

}} // namespace DocsExamples::Programming_with_Documents
