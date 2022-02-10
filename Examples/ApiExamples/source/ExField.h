#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Bookmark.h>
#include <Aspose.Words.Cpp/BookmarkCollection.h>
#include <Aspose.Words.Cpp/BookmarkEnd.h>
#include <Aspose.Words.Cpp/BookmarkStart.h>
#include <Aspose.Words.Cpp/Border.h>
#include <Aspose.Words.Cpp/BorderCollection.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/BuildingBlocks/BuildingBlock.h>
#include <Aspose.Words.Cpp/BuildingBlocks/BuildingBlockBehavior.h>
#include <Aspose.Words.Cpp/BuildingBlocks/BuildingBlockGallery.h>
#include <Aspose.Words.Cpp/BuildingBlocks/GlossaryDocument.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/ControlChar.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Fields/ComparisonEvaluationResult.h>
#include <Aspose.Words.Cpp/Fields/ComparisonExpression.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldAddIn.h>
#include <Aspose.Words.Cpp/Fields/FieldAddressBlock.h>
#include <Aspose.Words.Cpp/Fields/FieldAdvance.h>
#include <Aspose.Words.Cpp/Fields/FieldArgumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/FieldAuthor.h>
#include <Aspose.Words.Cpp/Fields/FieldAutoNum.h>
#include <Aspose.Words.Cpp/Fields/FieldAutoNumLgl.h>
#include <Aspose.Words.Cpp/Fields/FieldAutoNumOut.h>
#include <Aspose.Words.Cpp/Fields/FieldAutoText.h>
#include <Aspose.Words.Cpp/Fields/FieldAutoTextList.h>
#include <Aspose.Words.Cpp/Fields/FieldBarcode.h>
#include <Aspose.Words.Cpp/Fields/FieldBibliography.h>
#include <Aspose.Words.Cpp/Fields/FieldBidiOutline.h>
#include <Aspose.Words.Cpp/Fields/FieldBuilder.h>
#include <Aspose.Words.Cpp/Fields/FieldChar.h>
#include <Aspose.Words.Cpp/Fields/FieldCitation.h>
#include <Aspose.Words.Cpp/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Fields/FieldComments.h>
#include <Aspose.Words.Cpp/Fields/FieldCompare.h>
#include <Aspose.Words.Cpp/Fields/FieldCreateDate.h>
#include <Aspose.Words.Cpp/Fields/FieldData.h>
#include <Aspose.Words.Cpp/Fields/FieldDate.h>
#include <Aspose.Words.Cpp/Fields/FieldDde.h>
#include <Aspose.Words.Cpp/Fields/FieldDdeAuto.h>
#include <Aspose.Words.Cpp/Fields/FieldDisplayBarcode.h>
#include <Aspose.Words.Cpp/Fields/FieldDocProperty.h>
#include <Aspose.Words.Cpp/Fields/FieldDocVariable.h>
#include <Aspose.Words.Cpp/Fields/FieldEQ.h>
#include <Aspose.Words.Cpp/Fields/FieldEditTime.h>
#include <Aspose.Words.Cpp/Fields/FieldEnd.h>
#include <Aspose.Words.Cpp/Fields/FieldFileSize.h>
#include <Aspose.Words.Cpp/Fields/FieldFillIn.h>
#include <Aspose.Words.Cpp/Fields/FieldFootnoteRef.h>
#include <Aspose.Words.Cpp/Fields/FieldFormCheckBox.h>
#include <Aspose.Words.Cpp/Fields/FieldFormDropDown.h>
#include <Aspose.Words.Cpp/Fields/FieldFormText.h>
#include <Aspose.Words.Cpp/Fields/FieldFormat.h>
#include <Aspose.Words.Cpp/Fields/FieldFormula.h>
#include <Aspose.Words.Cpp/Fields/FieldGlossary.h>
#include <Aspose.Words.Cpp/Fields/FieldGoToButton.h>
#include <Aspose.Words.Cpp/Fields/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Fields/FieldIf.h>
#include <Aspose.Words.Cpp/Fields/FieldIfComparisonResult.h>
#include <Aspose.Words.Cpp/Fields/FieldImport.h>
#include <Aspose.Words.Cpp/Fields/FieldInclude.h>
#include <Aspose.Words.Cpp/Fields/FieldIncludePicture.h>
#include <Aspose.Words.Cpp/Fields/FieldIncludeText.h>
#include <Aspose.Words.Cpp/Fields/FieldIndex.h>
#include <Aspose.Words.Cpp/Fields/FieldIndexFormat.h>
#include <Aspose.Words.Cpp/Fields/FieldInfo.h>
#include <Aspose.Words.Cpp/Fields/FieldKeywords.h>
#include <Aspose.Words.Cpp/Fields/FieldLastSavedBy.h>
#include <Aspose.Words.Cpp/Fields/FieldLink.h>
#include <Aspose.Words.Cpp/Fields/FieldListNum.h>
#include <Aspose.Words.Cpp/Fields/FieldMacroButton.h>
#include <Aspose.Words.Cpp/Fields/FieldMergeField.h>
#include <Aspose.Words.Cpp/Fields/FieldNoteRef.h>
#include <Aspose.Words.Cpp/Fields/FieldNumChars.h>
#include <Aspose.Words.Cpp/Fields/FieldNumPages.h>
#include <Aspose.Words.Cpp/Fields/FieldNumWords.h>
#include <Aspose.Words.Cpp/Fields/FieldOcx.h>
#include <Aspose.Words.Cpp/Fields/FieldOptions.h>
#include <Aspose.Words.Cpp/Fields/FieldPage.h>
#include <Aspose.Words.Cpp/Fields/FieldPageRef.h>
#include <Aspose.Words.Cpp/Fields/FieldPrint.h>
#include <Aspose.Words.Cpp/Fields/FieldPrintDate.h>
#include <Aspose.Words.Cpp/Fields/FieldPrivate.h>
#include <Aspose.Words.Cpp/Fields/FieldQuote.h>
#include <Aspose.Words.Cpp/Fields/FieldRD.h>
#include <Aspose.Words.Cpp/Fields/FieldRef.h>
#include <Aspose.Words.Cpp/Fields/FieldRevNum.h>
#include <Aspose.Words.Cpp/Fields/FieldSaveDate.h>
#include <Aspose.Words.Cpp/Fields/FieldSection.h>
#include <Aspose.Words.Cpp/Fields/FieldSectionPages.h>
#include <Aspose.Words.Cpp/Fields/FieldSeparator.h>
#include <Aspose.Words.Cpp/Fields/FieldSeq.h>
#include <Aspose.Words.Cpp/Fields/FieldSet.h>
#include <Aspose.Words.Cpp/Fields/FieldStart.h>
#include <Aspose.Words.Cpp/Fields/FieldStyleRef.h>
#include <Aspose.Words.Cpp/Fields/FieldSubject.h>
#include <Aspose.Words.Cpp/Fields/FieldSymbol.h>
#include <Aspose.Words.Cpp/Fields/FieldTA.h>
#include <Aspose.Words.Cpp/Fields/FieldTC.h>
#include <Aspose.Words.Cpp/Fields/FieldTemplate.h>
#include <Aspose.Words.Cpp/Fields/FieldTime.h>
#include <Aspose.Words.Cpp/Fields/FieldTitle.h>
#include <Aspose.Words.Cpp/Fields/FieldToa.h>
#include <Aspose.Words.Cpp/Fields/FieldToc.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Fields/FieldUnknown.h>
#include <Aspose.Words.Cpp/Fields/FieldUpdateCultureSource.h>
#include <Aspose.Words.Cpp/Fields/FieldUserAddress.h>
#include <Aspose.Words.Cpp/Fields/FieldUserInitials.h>
#include <Aspose.Words.Cpp/Fields/FieldUserName.h>
#include <Aspose.Words.Cpp/Fields/FieldXE.h>
#include <Aspose.Words.Cpp/Fields/GeneralFormat.h>
#include <Aspose.Words.Cpp/Fields/GeneralFormatCollection.h>
#include <Aspose.Words.Cpp/Fields/IComparisonExpressionEvaluator.h>
#include <Aspose.Words.Cpp/Fields/IFieldUpdatingCallback.h>
#include <Aspose.Words.Cpp/Fields/IFieldUserPromptRespondent.h>
#include <Aspose.Words.Cpp/Fields/UserInformation.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Lists/List.h>
#include <Aspose.Words.Cpp/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Notes/Footnote.h>
#include <Aspose.Words.Cpp/Notes/FootnoteType.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Replacing/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Replacing/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Replacing/ReplaceAction.h>
#include <Aspose.Words.Cpp/Replacing/ReplacingArgs.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/Shading.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/VariableCollection.h>
#include <Aspose.Words.Cpp/VisitorAction.h>
#include <drawing/color.h>
#include <drawing/color_translator.h>
#include <net/http_status_code.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ilist.h>
#include <system/collections/list.h>
#include <system/convert.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/globalization/calendar.h>
#include <system/globalization/culture_info.h>
#include <system/globalization/date_time_format_info.h>
#include <system/globalization/um_al_qura_calendar.h>
#include <system/io/file.h>
#include <system/io/memory_stream.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/primitive_types.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/string_builder.h>
#include <system/timespan.h>
#include <system/timezone_info.h>
#include <xml/xml_attribute_collection.h>
#include <xml/xml_document.h>
#include <xml/xml_name_table.h>
#include <xml/xml_namespace_manager.h>
#include <xml/xml_node.h>
#include <xml/xml_node_list.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;

namespace ApiExamples {

class ExField : public ApiExampleBase
{
public:
    void GetFieldFromDocument()
    {
        //ExStart
        //ExFor:FieldType
        //ExFor:FieldChar
        //ExFor:FieldChar.FieldType
        //ExFor:FieldChar.IsDirty
        //ExFor:FieldChar.IsLocked
        //ExFor:FieldChar.GetField
        //ExFor:Field.IsLocked
        //ExSummary:Shows how to work with a FieldStart node.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto field = System::DynamicCast<FieldDate>(builder->InsertField(FieldType::FieldDate, true));
        field->get_Format()->set_DateTimeFormat(u"dddd, MMMM dd, yyyy");
        field->Update();

        SharedPtr<FieldChar> fieldStart = field->get_Start();

        ASSERT_EQ(FieldType::FieldDate, fieldStart->get_FieldType());
        ASPOSE_ASSERT_EQ(false, fieldStart->get_IsDirty());
        ASPOSE_ASSERT_EQ(false, fieldStart->get_IsLocked());

        // Retrieve the facade object which represents the field in the document.
        field = System::DynamicCast<FieldDate>(fieldStart->GetField());

        ASPOSE_ASSERT_EQ(false, field->get_IsLocked());
        ASSERT_EQ(u" DATE  \\@ \"dddd, MMMM dd, yyyy\"", field->GetFieldCode());

        // Update the field to show the current date.
        field->Update();
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);

        TestUtil::VerifyField(FieldType::FieldDate, u" DATE  \\@ \"dddd, MMMM dd, yyyy\"", System::DateTime::get_Now().ToString(u"dddd, MMMM dd, yyyy"),
                              doc->get_Range()->get_Fields()->idx_get(0));
    }

    void GetFieldCode()
    {
        //ExStart
        //ExFor:Field.GetFieldCode()
        //ExFor:Field.GetFieldCode(bool)
        //ExSummary:Shows how to get a field's field code.
        // Open a document which contains a MERGEFIELD inside an IF field.
        auto doc = MakeObject<Document>(MyDir + u"Nested fields.docx");
        auto fieldIf = System::DynamicCast<FieldIf>(doc->get_Range()->get_Fields()->idx_get(0));

        // There are two ways of getting a field's field code:
        // 1 -  Omit its inner fields:
        ASSERT_EQ(u" IF  > 0 \" (surplus of ) \" \"\" ", fieldIf->GetFieldCode(false));

        // 2 -  Include its inner fields:
        ASSERT_EQ(String::Format(u" IF \u0013 MERGEFIELD NetIncome \u0014\u0015 > 0 \" (surplus of \u0013 MERGEFIELD  NetIncome \\f $ \u0014\u0015) \" \"\" "),
                  fieldIf->GetFieldCode(true));

        // By default, the GetFieldCode method displays inner fields.
        ASSERT_EQ(fieldIf->GetFieldCode(), fieldIf->GetFieldCode(true));
        //ExEnd
    }

    void DisplayResult()
    {
        //ExStart
        //ExFor:Field.DisplayResult
        //ExSummary:Shows how to get the real text that a field displays in the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"This document was written by ");
        auto fieldAuthor = System::DynamicCast<FieldAuthor>(builder->InsertField(FieldType::FieldAuthor, true));
        fieldAuthor->set_AuthorName(u"John Doe");

        // We can use the DisplayResult property to verify what exact text
        // a field would display in its place in the document.
        ASSERT_EQ(String::Empty, fieldAuthor->get_DisplayResult());

        // Fields do not maintain accurate result values in real-time.
        // To make sure our fields display accurate results at any given time,
        // such as right before a save operation, we need to update them manually.
        fieldAuthor->Update();

        ASSERT_EQ(u"John Doe", fieldAuthor->get_DisplayResult());

        doc->Save(ArtifactsDir + u"Field.DisplayResult.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.DisplayResult.docx");

        ASSERT_EQ(u"John Doe", doc->get_Range()->get_Fields()->idx_get(0)->get_DisplayResult());
    }

    void CreateWithFieldBuilder()
    {
        //ExStart
        //ExFor:FieldBuilder.#ctor(FieldType)
        //ExFor:FieldBuilder.BuildAndInsert(Inline)
        //ExSummary:Shows how to create and insert a field using a field builder.
        auto doc = MakeObject<Document>();

        // A convenient way of adding text content to a document is with a document builder.
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u" Hello world! This text is one Run, which is an inline node.");

        // Fields have their builder, which we can use to construct a field code piece by piece.
        // In this case, we will construct a BARCODE field representing a US postal code,
        // and then insert it in front of a Run.
        auto fieldBuilder = MakeObject<FieldBuilder>(FieldType::FieldBarcode);
        fieldBuilder->AddArgument(u"90210");
        fieldBuilder->AddSwitch(u"\\f", u"A");
        fieldBuilder->AddSwitch(u"\\u");

        fieldBuilder->BuildAndInsert(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0));

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.CreateWithFieldBuilder.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.CreateWithFieldBuilder.docx");

        TestUtil::VerifyField(FieldType::FieldBarcode, u" BARCODE 90210 \\f A \\u ", String::Empty, doc->get_Range()->get_Fields()->idx_get(0));

        ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(11)->get_PreviousSibling(),
                         doc->get_Range()->get_Fields()->idx_get(0)->get_End());
        ASSERT_EQ(String::Format(u"{0} BARCODE 90210 \\f A \\u {1} Hello world! This text is one Run, which is an inline node.", ControlChar::FieldStartChar,
                                 ControlChar::FieldEndChar),
                  doc->GetText().Trim());
    }

    void RevNum()
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.RevisionNumber
        //ExFor:FieldRevNum
        //ExSummary:Shows how to work with REVNUM fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Current revision #");

        // Insert a REVNUM field, which displays the document's current revision number property.
        auto field = System::DynamicCast<FieldRevNum>(builder->InsertField(FieldType::FieldRevisionNum, true));

        ASSERT_EQ(u" REVNUM ", field->GetFieldCode());
        ASSERT_EQ(u"1", field->get_Result());
        ASSERT_EQ(1, doc->get_BuiltInDocumentProperties()->get_RevisionNumber());

        // This property counts how many times a document has been saved in Microsoft Word,
        // and is unrelated to tracked revisions. We can find it by right clicking the document in Windows Explorer
        // via Properties -> Details. We can update this property manually.
        doc->get_BuiltInDocumentProperties()->set_RevisionNumber(doc->get_BuiltInDocumentProperties()->get_RevisionNumber() + 1);
        ASSERT_EQ(u"1", field->get_Result());
        //ExSkip
        field->Update();

        ASSERT_EQ(u"2", field->get_Result());
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        ASSERT_EQ(2, doc->get_BuiltInDocumentProperties()->get_RevisionNumber());

        TestUtil::VerifyField(FieldType::FieldRevisionNum, u" REVNUM ", u"2", doc->get_Range()->get_Fields()->idx_get(0));
    }

    void InsertFieldNone()
    {
        //ExStart
        //ExFor:FieldUnknown
        //ExSummary:Shows how to work with 'FieldNone' field in a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a field that does not denote an objective field type in its field code.
        SharedPtr<Field> field = builder->InsertField(u" NOTAREALFIELD //a");

        // The "FieldNone" field type is reserved for fields such as these.
        ASSERT_EQ(FieldType::FieldNone, field->get_Type());

        // We can also still work with these fields and assign them as instances of the FieldUnknown class.
        auto fieldUnknown = System::DynamicCast<FieldUnknown>(field);
        ASSERT_EQ(u" NOTAREALFIELD //a", fieldUnknown->GetFieldCode());
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);

        TestUtil::VerifyField(FieldType::FieldNone, u" NOTAREALFIELD //a", u"Error! Bookmark not defined.", doc->get_Range()->get_Fields()->idx_get(0));
    }

    void InsertTcField()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a TC field at the current document builder position.
        builder->InsertField(u"TC \"Entry Text\" \\f t");
    }

    void InsertTcFieldsAtText()
    {
        auto doc = MakeObject<Document>();

        auto options = MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(MakeObject<ExField::InsertTcFieldHandler>(u"Chapter 1", u"\\l 1"));

        // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"The Beginning"), u"", options);
    }

    class InsertTcFieldHandler : public IReplacingCallback
    {
    public:
        /// <summary>
        /// The display text and switches to use for each TC field. Display name can be an empty String or null.
        /// </summary>
        InsertTcFieldHandler(String text, String switches)
        {
            mFieldText = text;
            mFieldSwitches = switches;
        }

    private:
        String mFieldText;
        String mFieldSwitches;

        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            auto builder = MakeObject<DocumentBuilder>(System::DynamicCast<Document>(args->get_MatchNode()->get_Document()));
            builder->MoveTo(args->get_MatchNode());

            // If the user-specified text is used in the field as display text, use that, otherwise
            // use the match String as the display text.
            String insertText = !String::IsNullOrEmpty(mFieldText) ? mFieldText : args->get_Match()->get_Value();

            // Insert the TC field before this node using the specified String
            // as the display text and user-defined switches.
            builder->InsertField(String::Format(u"TC \"{0}\" {1}", insertText, mFieldSwitches));

            return ReplaceAction::Skip;
        }
    };

    void FieldLocale()
    {
        //ExStart
        //ExFor:Field.LocaleId
        //ExSummary:Shows how to insert a field and work with its locale.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a DATE field, and then print the date it will display.
        // Your thread's current culture determines the formatting of the date.
        SharedPtr<Field> field = builder->InsertField(u"DATE");
        std::cout << "Today's date, as displayed in the \"" << System::Globalization::CultureInfo::get_CurrentCulture()->get_EnglishName()
                  << "\" culture: " << field->get_Result() << std::endl;

        ASSERT_EQ(1033, field->get_LocaleId());
        ASSERT_EQ(FieldUpdateCultureSource::CurrentThread, doc->get_FieldOptions()->get_FieldUpdateCultureSource());
        //ExSkip

        // Changing the culture of our thread will impact the result of the DATE field.
        // Another way to get the DATE field to display a date in a different culture is to use its LocaleId property.
        // This way allows us to avoid changing the thread's culture to get this effect.
        doc->get_FieldOptions()->set_FieldUpdateCultureSource(FieldUpdateCultureSource::FieldCode);
        auto de = MakeObject<System::Globalization::CultureInfo>(u"de-DE");
        field->set_LocaleId(de->get_LCID());
        field->Update();

        std::cout << "Today's date, as displayed according to the \""
                  << System::Globalization::CultureInfo::GetCultureInfo(field->get_LocaleId())->get_EnglishName() << "\" culture: " << field->get_Result()
                  << std::endl;
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        field = doc->get_Range()->get_Fields()->idx_get(0);

        TestUtil::VerifyField(FieldType::FieldDate, u"DATE", System::DateTime::get_Now().ToString(de->get_DateTimeFormat()->get_ShortDatePattern()), field);
        ASSERT_EQ(MakeObject<System::Globalization::CultureInfo>(u"de-DE")->get_LCID(), field->get_LocaleId());
    }

    void UpdateDirtyFields(bool updateDirtyFields)
    {
        //ExStart
        //ExFor:Field.IsDirty
        //ExFor:LoadOptions.UpdateDirtyFields
        //ExSummary:Shows how to use special property for updating field result.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Give the document's built-in "Author" property value, and then display it with a field.
        doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
        auto field = System::DynamicCast<FieldAuthor>(builder->InsertField(FieldType::FieldAuthor, true));

        ASSERT_FALSE(field->get_IsDirty());
        ASSERT_EQ(u"John Doe", field->get_Result());

        // Update the property. The field still displays the old value.
        doc->get_BuiltInDocumentProperties()->set_Author(u"John & Jane Doe");

        ASSERT_EQ(u"John Doe", field->get_Result());

        // Since the field's value is out of date, we can mark it as "dirty".
        // This value will stay out of date until we update the field manually with the Field.Update() method.
        field->set_IsDirty(true);

        {
            auto docStream = MakeObject<System::IO::MemoryStream>();
            // If we save without calling an update method,
            // the field will keep displaying the out of date value in the output document.
            doc->Save(docStream, SaveFormat::Docx);

            // The LoadOptions object has an option to update all fields
            // marked as "dirty" when loading the document.
            auto options = MakeObject<Loading::LoadOptions>();
            options->set_UpdateDirtyFields(updateDirtyFields);
            doc = MakeObject<Document>(docStream, options);

            ASSERT_EQ(u"John & Jane Doe", doc->get_BuiltInDocumentProperties()->get_Author());

            field = System::DynamicCast<FieldAuthor>(doc->get_Range()->get_Fields()->idx_get(0));

            // Updating dirty fields like this automatically set their "IsDirty" flag to false.
            if (updateDirtyFields)
            {
                ASSERT_EQ(u"John & Jane Doe", field->get_Result());
                ASSERT_FALSE(field->get_IsDirty());
            }
            else
            {
                ASSERT_EQ(u"John Doe", field->get_Result());
                ASSERT_TRUE(field->get_IsDirty());
            }
        }
        //ExEnd
    }

    void InsertFieldWithFieldBuilderException()
    {
        auto doc = MakeObject<Document>();

        SharedPtr<Run> run = DocumentHelper::InsertNewRun(doc, u" Hello World!", 0);

        auto argumentBuilder = MakeObject<FieldArgumentBuilder>();
        argumentBuilder->AddField(MakeObject<FieldBuilder>(FieldType::FieldMergeField));
        argumentBuilder->AddNode(run);
        argumentBuilder->AddText(u"Text argument builder");

        auto fieldBuilder = MakeObject<FieldBuilder>(FieldType::FieldIncludeText);

        ASSERT_THROW(static_cast<std::function<SharedPtr<Field>()>>(
                         [&fieldBuilder, &argumentBuilder, &run]() -> SharedPtr<Field> {
                             return fieldBuilder->AddArgument(argumentBuilder)
                                 ->AddArgument(u"=")
                                 ->AddArgument(u"BestField")
                                 ->AddArgument(10)
                                 ->AddArgument(20.0)
                                 ->BuildAndInsert(run);
                         })(),
                     System::ArgumentException);
    }

    void PreserveIncludePicture(bool preserveIncludePictureField)
    {
        //ExStart
        //ExFor:Field.Update(bool)
        //ExFor:LoadOptions.PreserveIncludePictureField
        //ExSummary:Shows how to preserve or discard INCLUDEPICTURE fields when loading a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto includePicture = System::DynamicCast<FieldIncludePicture>(builder->InsertField(FieldType::FieldIncludePicture, true));
        includePicture->set_SourceFullName(ImageDir + u"Transparent background logo.png");
        includePicture->Update(true);

        {
            auto docStream = MakeObject<System::IO::MemoryStream>();
            doc->Save(docStream, MakeObject<OoxmlSaveOptions>(SaveFormat::Docx));

            // We can set a flag in a LoadOptions object to decide whether to convert all INCLUDEPICTURE fields
            // into image shapes when loading a document that contains them.
            auto loadOptions = MakeObject<Loading::LoadOptions>();
            loadOptions->set_PreserveIncludePictureField(preserveIncludePictureField);

            doc = MakeObject<Document>(docStream, loadOptions);

            if (preserveIncludePictureField)
            {
                ASSERT_TRUE(doc->get_Range()->get_Fields()->LINQ_Any([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldIncludePicture; }));

                doc->UpdateFields();
                doc->Save(ArtifactsDir + u"Field.PreserveIncludePicture.docx");
            }
            else
            {
                ASSERT_FALSE(doc->get_Range()->get_Fields()->LINQ_Any([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldIncludePicture; }));
            }
        }
        //ExEnd
    }

    void FieldFormat_()
    {
        //ExStart
        //ExFor:Field.Format
        //ExFor:Field.Update()
        //ExFor:FieldFormat
        //ExFor:FieldFormat.DateTimeFormat
        //ExFor:FieldFormat.NumericFormat
        //ExFor:FieldFormat.GeneralFormats
        //ExFor:GeneralFormat
        //ExFor:GeneralFormatCollection
        //ExFor:GeneralFormatCollection.Add(GeneralFormat)
        //ExFor:GeneralFormatCollection.Count
        //ExFor:GeneralFormatCollection.Item(Int32)
        //ExFor:GeneralFormatCollection.Remove(GeneralFormat)
        //ExFor:GeneralFormatCollection.RemoveAt(Int32)
        //ExFor:GeneralFormatCollection.GetEnumerator
        //ExSummary:Shows how to format field results.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to insert a field that displays a result with no format applied.
        SharedPtr<Field> field = builder->InsertField(u"= 2 + 3");

        ASSERT_EQ(u"= 2 + 3", field->GetFieldCode());
        ASSERT_EQ(u"5", field->get_Result());

        // We can apply a format to a field's result using the field's properties.
        // Below are three types of formats that we can apply to a field's result.
        // 1 -  Numeric format:
        SharedPtr<FieldFormat> format = field->get_Format();
        format->set_NumericFormat(u"$###.00");
        field->Update();

        ASSERT_EQ(u"= 2 + 3 \\# $###.00", field->GetFieldCode());
        ASSERT_EQ(u"$  5.00", field->get_Result());

        // 2 -  Date/time format:
        field = builder->InsertField(u"DATE");
        format = field->get_Format();
        format->set_DateTimeFormat(u"dddd, MMMM dd, yyyy");
        field->Update();

        ASSERT_EQ(u"DATE \\@ \"dddd, MMMM dd, yyyy\"", field->GetFieldCode());
        std::cout << "Today's date, in " << format->get_DateTimeFormat() << " format:\n\t" << field->get_Result() << std::endl;

        // 3 -  General format:
        field = builder->InsertField(u"= 25 + 33");
        format = field->get_Format();
        format->get_GeneralFormats()->Add(GeneralFormat::LowercaseRoman);
        format->get_GeneralFormats()->Add(GeneralFormat::Upper);
        field->Update();

        int index = 0;
        {
            SharedPtr<System::Collections::Generic::IEnumerator<GeneralFormat>> generalFormatEnumerator = format->get_GeneralFormats()->GetEnumerator();
            while (generalFormatEnumerator->MoveNext())
            {
                std::cout << String::Format(u"General format index {0}: {1}", index++, generalFormatEnumerator->get_Current()) << std::endl;
            }
        }

        ASSERT_EQ(u"= 25 + 33 \\* roman \\* Upper", field->GetFieldCode());
        ASSERT_EQ(u"LVIII", field->get_Result());
        ASSERT_EQ(2, format->get_GeneralFormats()->get_Count());
        ASSERT_EQ(GeneralFormat::LowercaseRoman, format->get_GeneralFormats()->idx_get(0));

        // We can remove our formats to revert the field's result to its original form.
        format->get_GeneralFormats()->Remove(GeneralFormat::LowercaseRoman);
        format->get_GeneralFormats()->RemoveAt(0);
        ASSERT_EQ(0, format->get_GeneralFormats()->get_Count());
        field->Update();

        ASSERT_EQ(u"= 25 + 33  ", field->GetFieldCode());
        ASSERT_EQ(u"58", field->get_Result());
        ASSERT_EQ(0, format->get_GeneralFormats()->get_Count());
        //ExEnd
    }

    void Unlink()
    {
        //ExStart
        //ExFor:Document.UnlinkFields
        //ExSummary:Shows how to unlink all fields in the document.
        auto doc = MakeObject<Document>(MyDir + u"Linked fields.docx");

        doc->UnlinkFields();
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        String paraWithFields = DocumentHelper::GetParagraphText(doc, 0);

        ASSERT_EQ(u"Fields.Docx   Элементы указателя не найдены.     1.\r", paraWithFields);
    }

    void UnlinkAllFieldsInRange()
    {
        //ExStart
        //ExFor:Range.UnlinkFields
        //ExSummary:Shows how to unlink all fields in a range.
        auto doc = MakeObject<Document>(MyDir + u"Linked fields.docx");

        auto newSection = System::DynamicCast<Section>(doc->get_Sections()->idx_get(0)->Clone(true));
        doc->get_Sections()->Add(newSection);

        doc->get_Sections()->idx_get(1)->get_Range()->UnlinkFields();
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        String secWithFields = DocumentHelper::GetSectionText(doc, 1);

        ASSERT_TRUE(secWithFields.Trim().EndsWith(u"Fields.Docx   Элементы указателя не найдены.     3.\rОшибка! Не указана последовательность.    Fields.Docx "
                                                  u"  Элементы указателя не найдены.     4."));
    }

    void UnlinkSingleField()
    {
        //ExStart
        //ExFor:Field.Unlink
        //ExSummary:Shows how to unlink a field.
        auto doc = MakeObject<Document>(MyDir + u"Linked fields.docx");
        doc->get_Range()->get_Fields()->idx_get(1)->Unlink();
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        String paraWithFields = DocumentHelper::GetParagraphText(doc, 0);

        ASSERT_TRUE(paraWithFields.Trim().EndsWith(
            u"FILENAME  \\* Caps  \\* MERGEFORMAT \u0014Fields.Docx\u0015   Элементы указателя не найдены.     \u0013 LISTNUM  LegalDefault \u0015"));
    }

    void UpdateTocPageNumbers()
    {
        auto doc = MakeObject<Document>(MyDir + u"Field sample - TOC.docx");

        SharedPtr<Node> startNode = DocumentHelper::GetParagraph(doc, 2);
        SharedPtr<Node> endNode;

        SharedPtr<NodeCollection> paragraphCollection = doc->GetChildNodes(NodeType::Paragraph, true);

        for (const auto& para : System::IterateOver(paragraphCollection->LINQ_OfType<SharedPtr<Paragraph>>()))
        {
            for (const auto& run : System::IterateOver(para->get_Runs()->LINQ_OfType<SharedPtr<Run>>()))
            {
                if (run->get_Text().Contains(ControlChar::PageBreak()))
                {
                    endNode = run;
                    break;
                }
            }
        }

        if (startNode != nullptr && endNode != nullptr)
        {
            RemoveSequence(startNode, endNode);

            startNode->Remove();
            endNode->Remove();
        }

        SharedPtr<NodeCollection> fStart = doc->GetChildNodes(NodeType::FieldStart, true);

        for (const auto& field : System::IterateOver(fStart->LINQ_OfType<SharedPtr<FieldStart>>()))
        {
            FieldType fType = field->get_FieldType();
            if (fType == FieldType::FieldTOC)
            {
                auto para = System::DynamicCast<Paragraph>(field->GetAncestor(NodeType::Paragraph));
                para->get_Range()->UpdateFields();
                break;
            }
        }

        doc->Save(ArtifactsDir + u"Field.UpdateTocPageNumbers.docx");
    }

    static void RemoveSequence(SharedPtr<Node> start, SharedPtr<Node> end)
    {
        SharedPtr<Node> curNode = start->NextPreOrder(start->get_Document());
        while (curNode != nullptr && !System::ObjectExt::Equals(curNode, end))
        {
            SharedPtr<Node> nextNode = curNode->NextPreOrder(start->get_Document());

            if (curNode->get_IsComposite())
            {
                auto curComposite = System::DynamicCast<CompositeNode>(curNode);
                if (!curComposite->GetChildNodes(NodeType::Any, true)->Contains(end) && !curComposite->GetChildNodes(NodeType::Any, true)->Contains(start))
                {
                    nextNode = curNode->get_NextSibling();
                    curNode->Remove();
                }
            }
            else
            {
                curNode->Remove();
            }

            curNode = nextNode;
        }
    }

    void FieldAdvance_()
    {
        //ExStart
        //ExFor:Fields.FieldAdvance
        //ExFor:Fields.FieldAdvance.DownOffset
        //ExFor:Fields.FieldAdvance.HorizontalPosition
        //ExFor:Fields.FieldAdvance.LeftOffset
        //ExFor:Fields.FieldAdvance.RightOffset
        //ExFor:Fields.FieldAdvance.UpOffset
        //ExFor:Fields.FieldAdvance.VerticalPosition
        //ExSummary:Shows how to insert an ADVANCE field, and edit its properties.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"This text is in its normal place.");

        // Below are two ways of using the ADVANCE field to adjust the position of text that follows it.
        // The effects of an ADVANCE field continue to be applied until the paragraph ends,
        // or another ADVANCE field updates the offset/coordinate values.
        // 1 -  Specify a directional offset:
        auto field = System::DynamicCast<FieldAdvance>(builder->InsertField(FieldType::FieldAdvance, true));
        ASSERT_EQ(FieldType::FieldAdvance, field->get_Type());
        //ExSkip
        ASSERT_EQ(u" ADVANCE ", field->GetFieldCode());
        //ExSkip
        field->set_RightOffset(u"5");
        field->set_UpOffset(u"5");

        ASSERT_EQ(u" ADVANCE  \\r 5 \\u 5", field->GetFieldCode());

        builder->Write(u"This text will be moved up and to the right.");

        field = System::DynamicCast<FieldAdvance>(builder->InsertField(FieldType::FieldAdvance, true));
        field->set_DownOffset(u"5");
        field->set_LeftOffset(u"100");

        ASSERT_EQ(u" ADVANCE  \\d 5 \\l 100", field->GetFieldCode());

        builder->Writeln(u"This text is moved down and to the left, overlapping the previous text.");

        // 2 -  Move text to a position specified by coordinates:
        field = System::DynamicCast<FieldAdvance>(builder->InsertField(FieldType::FieldAdvance, true));
        field->set_HorizontalPosition(u"-100");
        field->set_VerticalPosition(u"200");

        ASSERT_EQ(u" ADVANCE  \\x -100 \\y 200", field->GetFieldCode());

        builder->Write(u"This text is in a custom position.");

        doc->Save(ArtifactsDir + u"Field.ADVANCE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.ADVANCE.docx");

        field = System::DynamicCast<FieldAdvance>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldAdvance, u" ADVANCE  \\r 5 \\u 5", String::Empty, field);
        ASSERT_EQ(u"5", field->get_RightOffset());
        ASSERT_EQ(u"5", field->get_UpOffset());

        field = System::DynamicCast<FieldAdvance>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldAdvance, u" ADVANCE  \\d 5 \\l 100", String::Empty, field);
        ASSERT_EQ(u"5", field->get_DownOffset());
        ASSERT_EQ(u"100", field->get_LeftOffset());

        field = System::DynamicCast<FieldAdvance>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldAdvance, u" ADVANCE  \\x -100 \\y 200", String::Empty, field);
        ASSERT_EQ(u"-100", field->get_HorizontalPosition());
        ASSERT_EQ(u"200", field->get_VerticalPosition());
    }

    void FieldAddressBlock_()
    {
        //ExStart
        //ExFor:Fields.FieldAddressBlock.ExcludedCountryOrRegionName
        //ExFor:Fields.FieldAddressBlock.FormatAddressOnCountryOrRegion
        //ExFor:Fields.FieldAddressBlock.IncludeCountryOrRegionName
        //ExFor:Fields.FieldAddressBlock.LanguageId
        //ExFor:Fields.FieldAddressBlock.NameAndAddressFormat
        //ExSummary:Shows how to insert an ADDRESSBLOCK field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto field = System::DynamicCast<FieldAddressBlock>(builder->InsertField(FieldType::FieldAddressBlock, true));

        ASSERT_EQ(u" ADDRESSBLOCK ", field->GetFieldCode());

        // Setting this to "2" will include all countries and regions,
        // unless it is the one specified in the ExcludedCountryOrRegionName property.
        field->set_IncludeCountryOrRegionName(u"2");
        field->set_FormatAddressOnCountryOrRegion(true);
        field->set_ExcludedCountryOrRegionName(u"United States");
        field->set_NameAndAddressFormat(u"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>");

        // By default, this property will contain the language ID of the first character of the document.
        // We can set a different culture for the field to format the result with like this.
        field->set_LanguageId(System::Convert::ToString(MakeObject<System::Globalization::CultureInfo>(u"en-US")->get_LCID()));

        ASSERT_EQ(
            u" ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033",
            field->GetFieldCode());
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        field = System::DynamicCast<FieldAddressBlock>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(
            FieldType::FieldAddressBlock,
            u" ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033",
            u"«AddressBlock»", field);
        ASSERT_EQ(u"2", field->get_IncludeCountryOrRegionName());
        ASPOSE_ASSERT_EQ(true, field->get_FormatAddressOnCountryOrRegion());
        ASSERT_EQ(u"United States", field->get_ExcludedCountryOrRegionName());
        ASSERT_EQ(u"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>", field->get_NameAndAddressFormat());
        ASSERT_EQ(u"1033", field->get_LanguageId());
    }

    //ExStart
    //ExFor:FieldCollection
    //ExFor:FieldCollection.Count
    //ExFor:FieldCollection.GetEnumerator
    //ExFor:FieldStart
    //ExFor:FieldStart.Accept(DocumentVisitor)
    //ExFor:FieldSeparator
    //ExFor:FieldSeparator.Accept(DocumentVisitor)
    //ExFor:FieldEnd
    //ExFor:FieldEnd.Accept(DocumentVisitor)
    //ExFor:FieldEnd.HasSeparator
    //ExFor:Field.End
    //ExFor:Field.Separator
    //ExFor:Field.Start
    //ExSummary:Shows how to work with a collection of fields.
    void FieldCollection_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u" DATE \\@ \"dddd, d MMMM yyyy\" ");
        builder->InsertField(u" TIME ");
        builder->InsertField(u" REVNUM ");
        builder->InsertField(u" AUTHOR  \"John Doe\" ");
        builder->InsertField(u" SUBJECT \"My Subject\" ");
        builder->InsertField(u" QUOTE \"Hello world!\" ");
        doc->UpdateFields();

        SharedPtr<FieldCollection> fields = doc->get_Range()->get_Fields();

        ASSERT_EQ(6, fields->get_Count());

        // Iterate over the field collection, and print contents and type
        // of every field using a custom visitor implementation.
        auto fieldVisitor = MakeObject<ExField::FieldVisitor>();

        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Field>>> fieldEnumerator = fields->GetEnumerator();
            while (fieldEnumerator->MoveNext())
            {
                if (fieldEnumerator->get_Current() != nullptr)
                {
                    fieldEnumerator->get_Current()->get_Start()->Accept(fieldVisitor);
                    if (fieldEnumerator->get_Current()->get_Separator() != nullptr)
                    {
                        fieldEnumerator->get_Current()->get_Separator()->Accept(fieldVisitor);
                    }
                    fieldEnumerator->get_Current()->get_End()->Accept(fieldVisitor);
                }
                else
                {
                    std::cout << "There are no fields in the document." << std::endl;
                }
            }
        }

        std::cout << fieldVisitor->GetText() << std::endl;
        TestFieldCollection(fieldVisitor->GetText());
        //ExSkip
    }

    /// <summary>
    /// Document visitor implementation that prints field info.
    /// </summary>
    class FieldVisitor : public DocumentVisitor
    {
    public:
        FieldVisitor()
        {
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
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        VisitorAction VisitFieldStart(SharedPtr<FieldStart> fieldStart) override
        {
            mBuilder->AppendLine(String(u"Found field: ") + System::ObjectExt::ToString(fieldStart->get_FieldType()));
            mBuilder->AppendLine(String(u"\tField code: ") + fieldStart->GetField()->GetFieldCode());
            mBuilder->AppendLine(String(u"\tDisplayed as: ") + fieldStart->GetField()->get_Result());

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        VisitorAction VisitFieldSeparator(SharedPtr<FieldSeparator> fieldSeparator) override
        {
            mBuilder->AppendLine(String(u"\tFound separator: ") + fieldSeparator->GetText());

            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        VisitorAction VisitFieldEnd(SharedPtr<FieldEnd> fieldEnd) override
        {
            mBuilder->AppendLine(String(u"End of field: ") + System::ObjectExt::ToString(fieldEnd->get_FieldType()));

            return VisitorAction::Continue;
        }

    private:
        SharedPtr<System::Text::StringBuilder> mBuilder;
    };
    //ExEnd

    void TestFieldCollection(String fieldVisitorText)
    {
        ASSERT_TRUE(fieldVisitorText.Contains(u"Found field: FieldDate"));
        ASSERT_TRUE(fieldVisitorText.Contains(u"Found field: FieldTime"));
        ASSERT_TRUE(fieldVisitorText.Contains(u"Found field: FieldRevisionNum"));
        ASSERT_TRUE(fieldVisitorText.Contains(u"Found field: FieldAuthor"));
        ASSERT_TRUE(fieldVisitorText.Contains(u"Found field: FieldSubject"));
        ASSERT_TRUE(fieldVisitorText.Contains(u"Found field: FieldQuote"));
    }

    void RemoveFields()
    {
        //ExStart
        //ExFor:FieldCollection
        //ExFor:FieldCollection.Count
        //ExFor:FieldCollection.Clear
        //ExFor:FieldCollection.Item(Int32)
        //ExFor:FieldCollection.Remove(Field)
        //ExFor:FieldCollection.RemoveAt(Int32)
        //ExFor:Field.Remove
        //ExSummary:Shows how to remove fields from a field collection.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u" DATE \\@ \"dddd, d MMMM yyyy\" ");
        builder->InsertField(u" TIME ");
        builder->InsertField(u" REVNUM ");
        builder->InsertField(u" AUTHOR  \"John Doe\" ");
        builder->InsertField(u" SUBJECT \"My Subject\" ");
        builder->InsertField(u" QUOTE \"Hello world!\" ");
        doc->UpdateFields();

        SharedPtr<FieldCollection> fields = doc->get_Range()->get_Fields();

        ASSERT_EQ(6, fields->get_Count());

        // Below are four ways of removing fields from a field collection.
        // 1 -  Get a field to remove itself:
        fields->idx_get(0)->Remove();
        ASSERT_EQ(5, fields->get_Count());

        // 2 -  Get the collection to remove a field that we pass to its removal method:
        SharedPtr<Field> lastField = fields->idx_get(3);
        fields->Remove(lastField);
        ASSERT_EQ(4, fields->get_Count());

        // 3 -  Remove a field from a collection at an index:
        fields->RemoveAt(2);
        ASSERT_EQ(3, fields->get_Count());

        // 4 -  Remove all the fields from the collection at once:
        fields->Clear();
        ASSERT_EQ(0, fields->get_Count());
        //ExEnd
    }

    void FieldCompare_()
    {
        //ExStart
        //ExFor:FieldCompare
        //ExFor:FieldCompare.ComparisonOperator
        //ExFor:FieldCompare.LeftExpression
        //ExFor:FieldCompare.RightExpression
        //ExSummary:Shows how to compare expressions using a COMPARE field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto field = System::DynamicCast<FieldCompare>(builder->InsertField(FieldType::FieldCompare, true));
        field->set_LeftExpression(u"3");
        field->set_ComparisonOperator(u"<");
        field->set_RightExpression(u"2");
        field->Update();

        // The COMPARE field displays a "0" or a "1", depending on its statement's truth.
        // The result of this statement is false so that this field will display a "0".
        ASSERT_EQ(u" COMPARE  3 < 2", field->GetFieldCode());
        ASSERT_EQ(u"0", field->get_Result());

        builder->Writeln();

        field = System::DynamicCast<FieldCompare>(builder->InsertField(FieldType::FieldCompare, true));
        field->set_LeftExpression(u"5");
        field->set_ComparisonOperator(u"=");
        field->set_RightExpression(u"2 + 3");
        field->Update();

        // This field displays a "1" since the statement is true.
        ASSERT_EQ(u" COMPARE  5 = \"2 + 3\"", field->GetFieldCode());
        ASSERT_EQ(u"1", field->get_Result());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.COMPARE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.COMPARE.docx");

        field = System::DynamicCast<FieldCompare>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldCompare, u" COMPARE  3 < 2", u"0", field);
        ASSERT_EQ(u"3", field->get_LeftExpression());
        ASSERT_EQ(u"<", field->get_ComparisonOperator());
        ASSERT_EQ(u"2", field->get_RightExpression());

        field = System::DynamicCast<FieldCompare>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldCompare, u" COMPARE  5 = \"2 + 3\"", u"1", field);
        ASSERT_EQ(u"5", field->get_LeftExpression());
        ASSERT_EQ(u"=", field->get_ComparisonOperator());
        ASSERT_EQ(u"\"2 + 3\"", field->get_RightExpression());
    }

    void FieldIf_()
    {
        //ExStart
        //ExFor:FieldIf
        //ExFor:FieldIf.ComparisonOperator
        //ExFor:FieldIf.EvaluateCondition
        //ExFor:FieldIf.FalseText
        //ExFor:FieldIf.LeftExpression
        //ExFor:FieldIf.RightExpression
        //ExFor:FieldIf.TrueText
        //ExFor:FieldIfComparisonResult
        //ExSummary:Shows how to insert an IF field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Statement 1: ");
        auto field = System::DynamicCast<FieldIf>(builder->InsertField(FieldType::FieldIf, true));
        field->set_LeftExpression(u"0");
        field->set_ComparisonOperator(u"=");
        field->set_RightExpression(u"1");

        // The IF field will display a string from either its "TrueText" property,
        // or its "FalseText" property, depending on the truth of the statement that we have constructed.
        field->set_TrueText(u"True");
        field->set_FalseText(u"False");
        field->Update();

        // In this case, "0 = 1" is incorrect, so the displayed result will be "False".
        ASSERT_EQ(u" IF  0 = 1 True False", field->GetFieldCode());
        ASSERT_EQ(FieldIfComparisonResult::False, field->EvaluateCondition());
        ASSERT_EQ(u"False", field->get_Result());

        builder->Write(u"\nStatement 2: ");
        field = System::DynamicCast<FieldIf>(builder->InsertField(FieldType::FieldIf, true));
        field->set_LeftExpression(u"5");
        field->set_ComparisonOperator(u"=");
        field->set_RightExpression(u"2 + 3");
        field->set_TrueText(u"True");
        field->set_FalseText(u"False");
        field->Update();

        // This time the statement is correct, so the displayed result will be "True".
        ASSERT_EQ(u" IF  5 = \"2 + 3\" True False", field->GetFieldCode());
        ASSERT_EQ(FieldIfComparisonResult::True, field->EvaluateCondition());
        ASSERT_EQ(u"True", field->get_Result());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.IF.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.IF.docx");
        field = System::DynamicCast<FieldIf>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldIf, u" IF  0 = 1 True False", u"False", field);
        ASSERT_EQ(u"0", field->get_LeftExpression());
        ASSERT_EQ(u"=", field->get_ComparisonOperator());
        ASSERT_EQ(u"1", field->get_RightExpression());
        ASSERT_EQ(u"True", field->get_TrueText());
        ASSERT_EQ(u"False", field->get_FalseText());

        field = System::DynamicCast<FieldIf>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldIf, u" IF  5 = \"2 + 3\" True False", u"True", field);
        ASSERT_EQ(u"5", field->get_LeftExpression());
        ASSERT_EQ(u"=", field->get_ComparisonOperator());
        ASSERT_EQ(u"\"2 + 3\"", field->get_RightExpression());
        ASSERT_EQ(u"True", field->get_TrueText());
        ASSERT_EQ(u"False", field->get_FalseText());
    }

    void FieldAutoNum_()
    {
        //ExStart
        //ExFor:FieldAutoNum
        //ExFor:FieldAutoNum.SeparatorCharacter
        //ExSummary:Shows how to number paragraphs using autonum fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Each AUTONUM field displays the current value of a running count of AUTONUM fields,
        // allowing us to automatically number items like a numbered list.
        // This field will display a number "1.".
        auto field = System::DynamicCast<FieldAutoNum>(builder->InsertField(FieldType::FieldAutoNum, true));
        builder->Writeln(u"\tParagraph 1.");

        ASSERT_EQ(u" AUTONUM ", field->GetFieldCode());

        field = System::DynamicCast<FieldAutoNum>(builder->InsertField(FieldType::FieldAutoNum, true));
        builder->Writeln(u"\tParagraph 2.");

        // The separator character, which appears in the field result immediately after the number,is a full stop by default.
        // If we leave this property null, our second AUTONUM field will display "2." in the document.
        ASSERT_TRUE(field->get_SeparatorCharacter() == nullptr);

        // We can set this property to apply the first character of its string as the new separator character.
        // In this case, our AUTONUM field will now display "2:".
        field->set_SeparatorCharacter(u":");

        ASSERT_EQ(u" AUTONUM  \\s :", field->GetFieldCode());

        doc->Save(ArtifactsDir + u"Field.AUTONUM.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.AUTONUM.docx");

        TestUtil::VerifyField(FieldType::FieldAutoNum, u" AUTONUM ", String::Empty, doc->get_Range()->get_Fields()->idx_get(0));
        TestUtil::VerifyField(FieldType::FieldAutoNum, u" AUTONUM  \\s :", String::Empty, doc->get_Range()->get_Fields()->idx_get(1));
    }

    //ExStart
    //ExFor:FieldAutoNumLgl
    //ExFor:FieldAutoNumLgl.RemoveTrailingPeriod
    //ExFor:FieldAutoNumLgl.SeparatorCharacter
    //ExSummary:Shows how to organize a document using AUTONUMLGL fields.
    void FieldAutoNumLgl_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        const String fillerText =
            String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ") +
            u"\nUt enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ";

        // AUTONUMLGL fields display a number that increments at each AUTONUMLGL field within its current heading level.
        // These fields maintain a separate count for each heading level,
        // and each field also displays the AUTONUMLGL field counts for all heading levels below its own.
        // Changing the count for any heading level resets the counts for all levels above that level to 1.
        // This allows us to organize our document in the form of an outline list.
        // This is the first AUTONUMLGL field at a heading level of 1, displaying "1." in the document.
        InsertNumberedClause(builder, u"\tHeading 1", fillerText, StyleIdentifier::Heading1);

        // This is the second AUTONUMLGL field at a heading level of 1, so it will display "2.".
        InsertNumberedClause(builder, u"\tHeading 2", fillerText, StyleIdentifier::Heading1);

        // This is the first AUTONUMLGL field at a heading level of 2,
        // and the AUTONUMLGL count for the heading level below it is "2", so it will display "2.1.".
        InsertNumberedClause(builder, u"\tHeading 3", fillerText, StyleIdentifier::Heading2);

        // This is the first AUTONUMLGL field at a heading level of 3.
        // Working in the same way as the field above, it will display "2.1.1.".
        InsertNumberedClause(builder, u"\tHeading 4", fillerText, StyleIdentifier::Heading3);

        // This field is at a heading level of 2, and its respective AUTONUMLGL count is at 2, so the field will display "2.2.".
        InsertNumberedClause(builder, u"\tHeading 5", fillerText, StyleIdentifier::Heading2);

        // Incrementing the AUTONUMLGL count for a heading level below this one
        // has reset the count for this level so that this field will display "2.2.1.".
        InsertNumberedClause(builder, u"\tHeading 6", fillerText, StyleIdentifier::Heading3);

        auto isAutoNumLegal = [](SharedPtr<Field> f)
        {
            return f->get_Type() == FieldType::FieldAutoNumLegal;
        };

        for (auto field : System::IterateOver<FieldAutoNumLgl>(doc->get_Range()->get_Fields()->LINQ_Where(isAutoNumLegal)))
        {
            // The separator character, which appears in the field result immediately after the number,
            // is a full stop by default. If we leave this property null,
            // our last AUTONUMLGL field will display "2.2.1." in the document.
            ASSERT_TRUE(field->get_SeparatorCharacter() == nullptr);

            // Setting a custom separator character and removing the trailing period
            // will change that field's appearance from "2.2.1." to "2:2:1".
            // We will apply this to all the fields that we have created.
            field->set_SeparatorCharacter(u":");
            field->set_RemoveTrailingPeriod(true);
            ASSERT_EQ(u" AUTONUMLGL  \\s : \\e", field->GetFieldCode());
        }

        doc->Save(ArtifactsDir + u"Field.AUTONUMLGL.docx");
        TestFieldAutoNumLgl(doc);
        //ExSkip
    }

    /// <summary>
    /// Uses a document builder to insert a clause numbered by an AUTONUMLGL field.
    /// </summary>
    static void InsertNumberedClause(SharedPtr<DocumentBuilder> builder, String heading, String contents, StyleIdentifier headingStyle)
    {
        builder->InsertField(FieldType::FieldAutoNumLegal, true);
        builder->get_CurrentParagraph()->get_ParagraphFormat()->set_StyleIdentifier(headingStyle);
        builder->Writeln(heading);

        // This text will belong to the auto num legal field above it.
        // It will collapse when we click the arrow next to the corresponding AUTONUMLGL field in Microsoft Word.
        builder->get_CurrentParagraph()->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::BodyText);
        builder->Writeln(contents);
    }
    //ExEnd

    void TestFieldAutoNumLgl(SharedPtr<Document> doc)
    {
        doc = DocumentHelper::SaveOpen(doc);

        auto isAutoNumLegal = [](SharedPtr<Field> f)
        {
            return f->get_Type() == FieldType::FieldAutoNumLegal;
        };

        for (auto field : System::IterateOver<FieldAutoNumLgl>(doc->get_Range()->get_Fields()->LINQ_Where(isAutoNumLegal)))
        {
            TestUtil::VerifyField(FieldType::FieldAutoNumLegal, u" AUTONUMLGL  \\s : \\e", String::Empty, field);

            ASSERT_EQ(u":", field->get_SeparatorCharacter());
            ASSERT_TRUE(field->get_RemoveTrailingPeriod());
        }
    }

    void FieldAutoNumOut_()
    {
        //ExStart
        //ExFor:FieldAutoNumOut
        //ExSummary:Shows how to number paragraphs using AUTONUMOUT fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // AUTONUMOUT fields display a number that increments at each AUTONUMOUT field.
        // Unlike AUTONUM fields, AUTONUMOUT fields use the outline numbering scheme,
        // which we can define in Microsoft Word via Format -> Bullets & Numbering -> "Outline Numbered".
        // This allows us to automatically number items like a numbered list.
        // LISTNUM fields are a newer alternative to AUTONUMOUT fields.
        // This field will display "1.".
        builder->InsertField(FieldType::FieldAutoNumOutline, true);
        builder->Writeln(u"\tParagraph 1.");

        // This field will display "2.".
        builder->InsertField(FieldType::FieldAutoNumOutline, true);
        builder->Writeln(u"\tParagraph 2.");

        for (auto field : System::IterateOver<FieldAutoNumOut>(
                 doc->get_Range()->get_Fields()->LINQ_Where([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldAutoNumOutline; })))
        {
            ASSERT_EQ(u" AUTONUMOUT ", field->GetFieldCode());
        }

        doc->Save(ArtifactsDir + u"Field.AUTONUMOUT.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.AUTONUMOUT.docx");

        for (const auto& field : System::IterateOver(doc->get_Range()->get_Fields()))
        {
            TestUtil::VerifyField(FieldType::FieldAutoNumOutline, u" AUTONUMOUT ", String::Empty, field);
        }
    }

    void FieldAutoText_()
    {
        //ExStart
        //ExFor:Fields.FieldAutoText
        //ExFor:FieldAutoText.EntryName
        //ExFor:FieldOptions.BuiltInTemplatesPaths
        //ExFor:FieldGlossary
        //ExFor:FieldGlossary.EntryName
        //ExSummary:Shows how to display a building block with AUTOTEXT and GLOSSARY fields.
        auto doc = MakeObject<Document>();

        // Create a glossary document and add an AutoText building block to it.
        doc->set_GlossaryDocument(MakeObject<GlossaryDocument>());
        auto buildingBlock = MakeObject<BuildingBlock>(doc->get_GlossaryDocument());
        buildingBlock->set_Name(u"MyBlock");
        buildingBlock->set_Gallery(BuildingBlockGallery::AutoText);
        buildingBlock->set_Category(u"General");
        buildingBlock->set_Description(u"MyBlock description");
        buildingBlock->set_Behavior(BuildingBlockBehavior::Paragraph);
        doc->get_GlossaryDocument()->AppendChild(buildingBlock);

        // Create a source and add it as text to our building block.
        auto buildingBlockSource = MakeObject<Document>();
        auto buildingBlockSourceBuilder = MakeObject<DocumentBuilder>(buildingBlockSource);
        buildingBlockSourceBuilder->Writeln(u"Hello World!");

        SharedPtr<Node> buildingBlockContent = doc->get_GlossaryDocument()->ImportNode(buildingBlockSource->get_FirstSection(), true);
        buildingBlock->AppendChild(buildingBlockContent);

        // Set a file which contains parts that our document, or its attached template may not contain.
        doc->get_FieldOptions()->set_BuiltInTemplatesPaths(MakeArray<String>({MyDir + u"Busniess brochure.dotx"}));

        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are two ways to use fields to display the contents of our building block.
        // 1 -  Using an AUTOTEXT field:
        auto fieldAutoText = System::DynamicCast<FieldAutoText>(builder->InsertField(FieldType::FieldAutoText, true));
        fieldAutoText->set_EntryName(u"MyBlock");

        ASSERT_EQ(u" AUTOTEXT  MyBlock", fieldAutoText->GetFieldCode());

        // 2 -  Using a GLOSSARY field:
        auto fieldGlossary = System::DynamicCast<FieldGlossary>(builder->InsertField(FieldType::FieldGlossary, true));
        fieldGlossary->set_EntryName(u"MyBlock");

        ASSERT_EQ(u" GLOSSARY  MyBlock", fieldGlossary->GetFieldCode());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.AUTOTEXT.GLOSSARY.dotx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.AUTOTEXT.GLOSSARY.dotx");

        ASSERT_EQ(0, doc->get_FieldOptions()->get_BuiltInTemplatesPaths()->get_Length());

        fieldAutoText = System::DynamicCast<FieldAutoText>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldAutoText, u" AUTOTEXT  MyBlock", u"Hello World!\r", fieldAutoText);
        ASSERT_EQ(u"MyBlock", fieldAutoText->get_EntryName());

        fieldGlossary = System::DynamicCast<FieldGlossary>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldGlossary, u" GLOSSARY  MyBlock", u"Hello World!\r", fieldGlossary);
        ASSERT_EQ(u"MyBlock", fieldGlossary->get_EntryName());
    }

    //ExStart
    //ExFor:Fields.FieldAutoTextList
    //ExFor:Fields.FieldAutoTextList.EntryName
    //ExFor:Fields.FieldAutoTextList.ListStyle
    //ExFor:Fields.FieldAutoTextList.ScreenTip
    //ExSummary:Shows how to use an AUTOTEXTLIST field to select from a list of AutoText entries.
    void FieldAutoTextList_()
    {
        auto doc = MakeObject<Document>();

        // Create a glossary document and populate it with auto text entries.
        doc->set_GlossaryDocument(MakeObject<GlossaryDocument>());
        AppendAutoTextEntry(doc->get_GlossaryDocument(), u"AutoText 1", u"Contents of AutoText 1");
        AppendAutoTextEntry(doc->get_GlossaryDocument(), u"AutoText 2", u"Contents of AutoText 2");
        AppendAutoTextEntry(doc->get_GlossaryDocument(), u"AutoText 3", u"Contents of AutoText 3");

        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an AUTOTEXTLIST field and set the text that the field will display in Microsoft Word.
        // Set the text to prompt the user to right-click this field to select an AutoText building block,
        // whose contents the field will display.
        auto field = System::DynamicCast<FieldAutoTextList>(builder->InsertField(FieldType::FieldAutoTextList, true));
        field->set_EntryName(u"Right click here to select an AutoText block");
        field->set_ListStyle(u"Heading 1");
        field->set_ScreenTip(u"Hover tip text for AutoTextList goes here");

        ASSERT_EQ(String(u" AUTOTEXTLIST  \"Right click here to select an AutoText block\" ") + u"\\s \"Heading 1\" " +
                      u"\\t \"Hover tip text for AutoTextList goes here\"",
                  field->GetFieldCode());

        doc->Save(ArtifactsDir + u"Field.AUTOTEXTLIST.dotx");
    }

    /// <summary>
    /// Create an AutoText-type building block and add it to a glossary document.
    /// </summary>
    static void AppendAutoTextEntry(SharedPtr<GlossaryDocument> glossaryDoc, String name, String contents)
    {
        auto buildingBlock = MakeObject<BuildingBlock>(glossaryDoc);
        buildingBlock->set_Name(name);
        buildingBlock->set_Gallery(BuildingBlockGallery::AutoText);
        buildingBlock->set_Category(u"General");
        buildingBlock->set_Behavior(BuildingBlockBehavior::Paragraph);

        auto section = MakeObject<Section>(glossaryDoc);
        section->AppendChild(MakeObject<Body>(glossaryDoc));
        section->get_Body()->AppendParagraph(contents);
        buildingBlock->AppendChild(section);

        glossaryDoc->AppendChild(buildingBlock);
    }
    //ExEnd

    void FieldListNum_()
    {
        //ExStart
        //ExFor:FieldListNum
        //ExFor:FieldListNum.HasListName
        //ExFor:FieldListNum.ListLevel
        //ExFor:FieldListNum.ListName
        //ExFor:FieldListNum.StartingNumber
        //ExSummary:Shows how to number paragraphs with LISTNUM fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // LISTNUM fields display a number that increments at each LISTNUM field.
        // These fields also have a variety of options that allow us to use them to emulate numbered lists.
        auto field = System::DynamicCast<FieldListNum>(builder->InsertField(FieldType::FieldListNum, true));

        // Lists start counting at 1 by default, but we can set this number to a different value, such as 0.
        // This field will display "0)".
        field->set_StartingNumber(u"0");
        builder->Writeln(u"Paragraph 1");

        ASSERT_EQ(u" LISTNUM  \\s 0", field->GetFieldCode());

        // LISTNUM fields maintain separate counts for each list level.
        // Inserting a LISTNUM field in the same paragraph as another LISTNUM field
        // increases the list level instead of the count.
        // The next field will continue the count we started above and display a value of "1" at list level 1.
        builder->InsertField(FieldType::FieldListNum, true);

        // This field will start a count at list level 2. It will display a value of "1".
        builder->InsertField(FieldType::FieldListNum, true);

        // This field will start a count at list level 3. It will display a value of "1".
        // Different list levels have different formatting,
        // so these fields combined will display a value of "1)a)i)".
        builder->InsertField(FieldType::FieldListNum, true);
        builder->Writeln(u"Paragraph 2");

        // The next LISTNUM field that we insert will continue the count at the list level
        // that the previous LISTNUM field was on.
        // We can use the "ListLevel" property to jump to a different list level.
        // If this LISTNUM field stayed on list level 3, it would display "ii)",
        // but, since we have moved it to list level 2, it carries on the count at that level and displays "b)".
        field = System::DynamicCast<FieldListNum>(builder->InsertField(FieldType::FieldListNum, true));
        field->set_ListLevel(u"2");
        builder->Writeln(u"Paragraph 3");

        ASSERT_EQ(u" LISTNUM  \\l 2", field->GetFieldCode());

        // We can set the ListName property to get the field to emulate a different AUTONUM field type.
        // "NumberDefault" emulates AUTONUM, "OutlineDefault" emulates AUTONUMOUT,
        // and "LegalDefault" emulates AUTONUMLGL fields.
        // The "OutlineDefault" list name with 1 as the starting number will result in displaying "I.".
        field = System::DynamicCast<FieldListNum>(builder->InsertField(FieldType::FieldListNum, true));
        field->set_StartingNumber(u"1");
        field->set_ListName(u"OutlineDefault");
        builder->Writeln(u"Paragraph 4");

        ASSERT_TRUE(field->get_HasListName());
        ASSERT_EQ(u" LISTNUM  OutlineDefault \\s 1", field->GetFieldCode());

        // The ListName does not carry over from the previous field, so we will need to set it for each new field.
        // This field continues the count with the different list name and displays "II.".
        field = System::DynamicCast<FieldListNum>(builder->InsertField(FieldType::FieldListNum, true));
        field->set_ListName(u"OutlineDefault");
        builder->Writeln(u"Paragraph 5");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.LISTNUM.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.LISTNUM.docx");

        ASSERT_EQ(7, doc->get_Range()->get_Fields()->get_Count());

        field = System::DynamicCast<FieldListNum>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldListNum, u" LISTNUM  \\s 0", String::Empty, field);
        ASSERT_EQ(u"0", field->get_StartingNumber());
        ASSERT_TRUE(field->get_ListLevel() == nullptr);
        ASSERT_FALSE(field->get_HasListName());
        ASSERT_TRUE(field->get_ListName() == nullptr);

        for (int i = 1; i < 4; i++)
        {
            field = System::DynamicCast<FieldListNum>(doc->get_Range()->get_Fields()->idx_get(i));

            TestUtil::VerifyField(FieldType::FieldListNum, u" LISTNUM ", String::Empty, field);
            ASSERT_TRUE(field->get_StartingNumber() == nullptr);
            ASSERT_TRUE(field->get_ListLevel() == nullptr);
            ASSERT_FALSE(field->get_HasListName());
            ASSERT_TRUE(field->get_ListName() == nullptr);
        }

        field = System::DynamicCast<FieldListNum>(doc->get_Range()->get_Fields()->idx_get(4));

        TestUtil::VerifyField(FieldType::FieldListNum, u" LISTNUM  \\l 2", String::Empty, field);
        ASSERT_TRUE(field->get_StartingNumber() == nullptr);
        ASSERT_EQ(u"2", field->get_ListLevel());
        ASSERT_FALSE(field->get_HasListName());
        ASSERT_TRUE(field->get_ListName() == nullptr);

        field = System::DynamicCast<FieldListNum>(doc->get_Range()->get_Fields()->idx_get(5));

        TestUtil::VerifyField(FieldType::FieldListNum, u" LISTNUM  OutlineDefault \\s 1", String::Empty, field);
        ASSERT_EQ(u"1", field->get_StartingNumber());
        ASSERT_TRUE(field->get_ListLevel() == nullptr);
        ASSERT_TRUE(field->get_HasListName());
        ASSERT_EQ(u"OutlineDefault", field->get_ListName());
    }

    //ExStart
    //ExFor:FieldToc
    //ExFor:FieldToc.BookmarkName
    //ExFor:FieldToc.CustomStyles
    //ExFor:FieldToc.EntrySeparator
    //ExFor:FieldToc.HeadingLevelRange
    //ExFor:FieldToc.HideInWebLayout
    //ExFor:FieldToc.InsertHyperlinks
    //ExFor:FieldToc.PageNumberOmittingLevelRange
    //ExFor:FieldToc.PreserveLineBreaks
    //ExFor:FieldToc.PreserveTabs
    //ExFor:FieldToc.UpdatePageNumbers
    //ExFor:FieldToc.UseParagraphOutlineLevel
    //ExFor:FieldOptions.CustomTocStyleSeparator
    //ExSummary:Shows how to insert a TOC, and populate it with entries based on heading styles.
    void FieldToc_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartBookmark(u"MyBookmark");

        // Insert a TOC field, which will compile all headings into a table of contents.
        // For each heading, this field will create a line with the text in that heading style to the left,
        // and the page the heading appears on to the right.
        auto field = System::DynamicCast<FieldToc>(builder->InsertField(FieldType::FieldTOC, true));

        // Use the BookmarkName property to only list headings
        // that appear within the bounds of a bookmark with the "MyBookmark" name.
        field->set_BookmarkName(u"MyBookmark");

        // Text with a built-in heading style, such as "Heading 1", applied to it will count as a heading.
        // We can name additional styles to be picked up as headings by the TOC in this property and their TOC levels.
        field->set_CustomStyles(u"Quote; 6; Intense Quote; 7");

        // By default, Styles/TOC levels are separated in the CustomStyles property by a comma,
        // but we can set a custom delimiter in this property.
        doc->get_FieldOptions()->set_CustomTocStyleSeparator(u";");

        // Configure the field to exclude any headings that have TOC levels outside of this range.
        field->set_HeadingLevelRange(u"1-3");

        // The TOC will not display the page numbers of headings whose TOC levels are within this range.
        field->set_PageNumberOmittingLevelRange(u"2-5");

        // Set a custom string that will separate every heading from its page number.
        field->set_EntrySeparator(u"-");
        field->set_InsertHyperlinks(true);
        field->set_HideInWebLayout(false);
        field->set_PreserveLineBreaks(true);
        field->set_PreserveTabs(true);
        field->set_UseParagraphOutlineLevel(false);

        InsertNewPageWithHeading(builder, u"First entry", u"Heading 1");
        builder->Writeln(u"Paragraph text.");
        InsertNewPageWithHeading(builder, u"Second entry", u"Heading 1");
        InsertNewPageWithHeading(builder, u"Third entry", u"Quote");
        InsertNewPageWithHeading(builder, u"Fourth entry", u"Intense Quote");

        // These two headings will have the page numbers omitted because they are within the "2-5" range.
        InsertNewPageWithHeading(builder, u"Fifth entry", u"Heading 2");
        InsertNewPageWithHeading(builder, u"Sixth entry", u"Heading 3");

        // This entry does not appear because "Heading 4" is outside of the "1-3" range that we have set earlier.
        InsertNewPageWithHeading(builder, u"Seventh entry", u"Heading 4");

        builder->EndBookmark(u"MyBookmark");
        builder->Writeln(u"Paragraph text.");

        // This entry does not appear because it is outside the bookmark specified by the TOC.
        InsertNewPageWithHeading(builder, u"Eighth entry", u"Heading 1");

        ASSERT_EQ(u" TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w", field->GetFieldCode());

        field->UpdatePageNumbers();
        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.TOC.docx");
        TestFieldToc(doc);
        //ExSkip
    }

    /// <summary>
    /// Start a new page and insert a paragraph of a specified style.
    /// </summary>
    void InsertNewPageWithHeading(SharedPtr<DocumentBuilder> builder, String captionText, String styleName)
    {
        builder->InsertBreak(BreakType::PageBreak);
        String originalStyle = builder->get_ParagraphFormat()->get_StyleName();
        builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(styleName));
        builder->Writeln(captionText);
        builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(originalStyle));
    }
    //ExEnd

    void TestFieldToc(SharedPtr<Document> doc)
    {
        doc = DocumentHelper::SaveOpen(doc);
        auto field = System::DynamicCast<FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(u"MyBookmark", field->get_BookmarkName());
        ASSERT_EQ(u"Quote; 6; Intense Quote; 7", field->get_CustomStyles());
        ASSERT_EQ(u"-", field->get_EntrySeparator());
        ASSERT_EQ(u"1-3", field->get_HeadingLevelRange());
        ASSERT_EQ(u"2-5", field->get_PageNumberOmittingLevelRange());
        ASSERT_FALSE(field->get_HideInWebLayout());
        ASSERT_TRUE(field->get_InsertHyperlinks());
        ASSERT_TRUE(field->get_PreserveLineBreaks());
        ASSERT_TRUE(field->get_PreserveTabs());
        ASSERT_TRUE(field->UpdatePageNumbers());
        ASSERT_FALSE(field->get_UseParagraphOutlineLevel());
        ASSERT_EQ(u" TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w", field->GetFieldCode());
        ASSERT_EQ(String(u"\u0013 HYPERLINK \\l \"_Toc256000001\" \u0014First entry-\u0013 PAGEREF _Toc256000001 \\h \u00142\u0015\u0015\r") +
                      u"\u0013 HYPERLINK \\l \"_Toc256000002\" \u0014Second entry-\u0013 PAGEREF _Toc256000002 \\h \u00143\u0015\u0015\r" +
                      u"\u0013 HYPERLINK \\l \"_Toc256000003\" \u0014Third entry-\u0013 PAGEREF _Toc256000003 \\h \u00144\u0015\u0015\r" +
                      u"\u0013 HYPERLINK \\l \"_Toc256000004\" \u0014Fourth entry-\u0013 PAGEREF _Toc256000004 \\h \u00145\u0015\u0015\r" +
                      u"\u0013 HYPERLINK \\l \"_Toc256000005\" \u0014Fifth entry\u0015\r" + u"\u0013 HYPERLINK \\l \"_Toc256000006\" \u0014Sixth entry\u0015\r",
                  field->get_Result());
    }

    //ExStart
    //ExFor:FieldToc.EntryIdentifier
    //ExFor:FieldToc.EntryLevelRange
    //ExFor:FieldTC
    //ExFor:FieldTC.OmitPageNumber
    //ExFor:FieldTC.Text
    //ExFor:FieldTC.TypeIdentifier
    //ExFor:FieldTC.EntryLevel
    //ExSummary:Shows how to insert a TOC field, and filter which TC fields end up as entries.
    void FieldTocEntryIdentifier()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a TOC field, which will compile all TC fields into a table of contents.
        auto fieldToc = System::DynamicCast<FieldToc>(builder->InsertField(FieldType::FieldTOC, true));

        // Configure the field only to pick up TC entries of the "A" type, and an entry-level between 1 and 3.
        fieldToc->set_EntryIdentifier(u"A");
        fieldToc->set_EntryLevelRange(u"1-3");

        ASSERT_EQ(u" TOC  \\f A \\l 1-3", fieldToc->GetFieldCode());

        // These two entries will appear in the table.
        builder->InsertBreak(BreakType::PageBreak);
        InsertTocEntry(builder, u"TC field 1", u"A", u"1");
        InsertTocEntry(builder, u"TC field 2", u"A", u"2");

        ASSERT_EQ(u" TC  \"TC field 1\" \\n \\f A \\l 1", doc->get_Range()->get_Fields()->idx_get(1)->GetFieldCode());

        // This entry will be omitted from the table because it has a different type from "A".
        InsertTocEntry(builder, u"TC field 3", u"B", u"1");

        // This entry will be omitted from the table because it has an entry-level outside of the 1-3 range.
        InsertTocEntry(builder, u"TC field 4", u"A", u"5");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.TC.docx");
        TestFieldTocEntryIdentifier(doc);
        //ExSkip
    }

    /// <summary>
    /// Use a document builder to insert a TC field.
    /// </summary>
    void InsertTocEntry(SharedPtr<DocumentBuilder> builder, String text, String typeIdentifier, String entryLevel)
    {
        auto fieldTc = System::DynamicCast<FieldTC>(builder->InsertField(FieldType::FieldTOCEntry, true));
        fieldTc->set_OmitPageNumber(true);
        fieldTc->set_Text(text);
        fieldTc->set_TypeIdentifier(typeIdentifier);
        fieldTc->set_EntryLevel(entryLevel);
    }
    //ExEnd

    void TestFieldTocEntryIdentifier(SharedPtr<Document> doc)
    {
        doc = DocumentHelper::SaveOpen(doc);
        auto fieldToc = System::DynamicCast<FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldTOC, u" TOC  \\f A \\l 1-3", u"TC field 1\rTC field 2\r", fieldToc);
        ASSERT_EQ(u"A", fieldToc->get_EntryIdentifier());
        ASSERT_EQ(u"1-3", fieldToc->get_EntryLevelRange());

        auto fieldTc = System::DynamicCast<FieldTC>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldTOCEntry, u" TC  \"TC field 1\" \\n \\f A \\l 1", String::Empty, fieldTc);
        ASSERT_TRUE(fieldTc->get_OmitPageNumber());
        ASSERT_EQ(u"TC field 1", fieldTc->get_Text());
        ASSERT_EQ(u"A", fieldTc->get_TypeIdentifier());
        ASSERT_EQ(u"1", fieldTc->get_EntryLevel());

        fieldTc = System::DynamicCast<FieldTC>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldTOCEntry, u" TC  \"TC field 2\" \\n \\f A \\l 2", String::Empty, fieldTc);
        ASSERT_TRUE(fieldTc->get_OmitPageNumber());
        ASSERT_EQ(u"TC field 2", fieldTc->get_Text());
        ASSERT_EQ(u"A", fieldTc->get_TypeIdentifier());
        ASSERT_EQ(u"2", fieldTc->get_EntryLevel());

        fieldTc = System::DynamicCast<FieldTC>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldTOCEntry, u" TC  \"TC field 3\" \\n \\f B \\l 1", String::Empty, fieldTc);
        ASSERT_TRUE(fieldTc->get_OmitPageNumber());
        ASSERT_EQ(u"TC field 3", fieldTc->get_Text());
        ASSERT_EQ(u"B", fieldTc->get_TypeIdentifier());
        ASSERT_EQ(u"1", fieldTc->get_EntryLevel());

        fieldTc = System::DynamicCast<FieldTC>(doc->get_Range()->get_Fields()->idx_get(4));

        TestUtil::VerifyField(FieldType::FieldTOCEntry, u" TC  \"TC field 4\" \\n \\f A \\l 5", String::Empty, fieldTc);
        ASSERT_TRUE(fieldTc->get_OmitPageNumber());
        ASSERT_EQ(u"TC field 4", fieldTc->get_Text());
        ASSERT_EQ(u"A", fieldTc->get_TypeIdentifier());
        ASSERT_EQ(u"5", fieldTc->get_EntryLevel());
    }

    void TocSeqPrefix()
    {
        //ExStart
        //ExFor:FieldToc
        //ExFor:FieldToc.TableOfFiguresLabel
        //ExFor:FieldToc.PrefixedSequenceIdentifier
        //ExFor:FieldToc.SequenceSeparator
        //ExFor:FieldSeq
        //ExFor:FieldSeq.SequenceIdentifier
        //ExSummary:Shows how to populate a TOC field with entries using SEQ fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // A TOC field can create an entry in its table of contents for each SEQ field found in the document.
        // Each entry contains the paragraph that includes the SEQ field and the page's number that the field appears on.
        auto fieldToc = System::DynamicCast<FieldToc>(builder->InsertField(FieldType::FieldTOC, true));

        // SEQ fields display a count that increments at each SEQ field.
        // These fields also maintain separate counts for each unique named sequence
        // identified by the SEQ field's "SequenceIdentifier" property.
        // Use the "TableOfFiguresLabel" property to name a main sequence for the TOC.
        // Now, this TOC will only create entries out of SEQ fields with their "SequenceIdentifier" set to "MySequence".
        fieldToc->set_TableOfFiguresLabel(u"MySequence");

        // We can name another SEQ field sequence in the "PrefixedSequenceIdentifier" property.
        // SEQ fields from this prefix sequence will not create TOC entries.
        // Every TOC entry created from a main sequence SEQ field will now also display the count that
        // the prefix sequence is currently on at the primary sequence SEQ field that made the entry.
        fieldToc->set_PrefixedSequenceIdentifier(u"PrefixSequence");

        // Each TOC entry will display the prefix sequence count immediately to the left
        // of the page number that the main sequence SEQ field appears on.
        // We can specify a custom separator that will appear between these two numbers.
        fieldToc->set_SequenceSeparator(u">");

        ASSERT_EQ(u" TOC  \\c MySequence \\s PrefixSequence \\d >", fieldToc->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);

        // There are two ways of using SEQ fields to populate this TOC.
        // 1 -  Inserting a SEQ field that belongs to the TOC's prefix sequence:
        // This field will increment the SEQ sequence count for the "PrefixSequence" by 1.
        // Since this field does not belong to the main sequence identified
        // by the "TableOfFiguresLabel" property of the TOC, it will not appear as an entry.
        auto fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"PrefixSequence");
        builder->InsertParagraph();

        ASSERT_EQ(u" SEQ  PrefixSequence", fieldSeq->GetFieldCode());

        // 2 -  Inserting a SEQ field that belongs to the TOC's main sequence:
        // This SEQ field will create an entry in the TOC.
        // The TOC entry will contain the paragraph that the SEQ field is in and the number of the page that it appears on.
        // This entry will also display the count that the prefix sequence is currently at,
        // separated from the page number by the value in the TOC's SeqenceSeparator property.
        // The "PrefixSequence" count is at 1, this main sequence SEQ field is on page 2,
        // and the separator is ">", so entry will display "1>2".
        builder->Write(u"First TOC entry, MySequence #");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");

        ASSERT_EQ(u" SEQ  MySequence", fieldSeq->GetFieldCode());

        // Insert a page, advance the prefix sequence by 2, and insert a SEQ field to create a TOC entry afterwards.
        // The prefix sequence is now at 2, and the main sequence SEQ field is on page 3,
        // so the TOC entry will display "2>3" at its page count.
        builder->InsertBreak(BreakType::PageBreak);
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"PrefixSequence");
        builder->InsertParagraph();
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        builder->Write(u"Second TOC entry, MySequence #");
        fieldSeq->set_SequenceIdentifier(u"MySequence");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.TOC.SEQ.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.TOC.SEQ.docx");

        ASSERT_EQ(9, doc->get_Range()->get_Fields()->get_Count());

        fieldToc = System::DynamicCast<FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));
        std::cout << fieldToc->get_DisplayResult() << std::endl;
        TestUtil::VerifyField(FieldType::FieldTOC, u" TOC  \\c MySequence \\s PrefixSequence \\d >",
                              String(u"First TOC entry, MySequence #12\t\u0013 SEQ PrefixSequence _Toc256000000 \\* ARABIC \u00141\u0015>\u0013 PAGEREF "
                                     u"_Toc256000000 \\h \u00142\u0015\r2") +
                                  u"Second TOC entry, MySequence #\t\u0013 SEQ PrefixSequence _Toc256000001 \\* ARABIC \u00142\u0015>\u0013 PAGEREF "
                                  u"_Toc256000001 \\h \u00143\u0015\r",
                              fieldToc);
        ASSERT_EQ(u"MySequence", fieldToc->get_TableOfFiguresLabel());
        ASSERT_EQ(u"PrefixSequence", fieldToc->get_PrefixedSequenceIdentifier());
        ASSERT_EQ(u">", fieldToc->get_SequenceSeparator());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ PrefixSequence _Toc256000000 \\* ARABIC ", u"1", fieldSeq);
        ASSERT_EQ(u"PrefixSequence", fieldSeq->get_SequenceIdentifier());

        // Byproduct field created by Aspose.Words
        auto fieldPageRef = System::DynamicCast<FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF _Toc256000000 \\h ", u"2", fieldPageRef);
        ASSERT_EQ(u"PrefixSequence", fieldSeq->get_SequenceIdentifier());
        ASSERT_EQ(u"_Toc256000000", fieldPageRef->get_BookmarkName());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ PrefixSequence _Toc256000001 \\* ARABIC ", u"2", fieldSeq);
        ASSERT_EQ(u"PrefixSequence", fieldSeq->get_SequenceIdentifier());

        fieldPageRef = System::DynamicCast<FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(4));

        TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF _Toc256000001 \\h ", u"3", fieldPageRef);
        ASSERT_EQ(u"PrefixSequence", fieldSeq->get_SequenceIdentifier());
        ASSERT_EQ(u"_Toc256000001", fieldPageRef->get_BookmarkName());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(5));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  PrefixSequence", u"1", fieldSeq);
        ASSERT_EQ(u"PrefixSequence", fieldSeq->get_SequenceIdentifier());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(6));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"1", fieldSeq);
        ASSERT_EQ(u"MySequence", fieldSeq->get_SequenceIdentifier());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(7));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  PrefixSequence", u"2", fieldSeq);
        ASSERT_EQ(u"PrefixSequence", fieldSeq->get_SequenceIdentifier());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(8));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"2", fieldSeq);
        ASSERT_EQ(u"MySequence", fieldSeq->get_SequenceIdentifier());
    }

    void TocSeqNumbering()
    {
        //ExStart
        //ExFor:FieldSeq
        //ExFor:FieldSeq.InsertNextNumber
        //ExFor:FieldSeq.ResetHeadingLevel
        //ExFor:FieldSeq.ResetNumber
        //ExFor:FieldSeq.SequenceIdentifier
        //ExSummary:Shows create numbering using SEQ fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // SEQ fields display a count that increments at each SEQ field.
        // These fields also maintain separate counts for each unique named sequence
        // identified by the SEQ field's "SequenceIdentifier" property.
        // Insert a SEQ field that will display the current count value of "MySequence",
        // after using the "ResetNumber" property to set it to 100.
        builder->Write(u"#");
        auto fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        fieldSeq->set_ResetNumber(u"100");
        fieldSeq->Update();

        ASSERT_EQ(u" SEQ  MySequence \\r 100", fieldSeq->GetFieldCode());
        ASSERT_EQ(u"100", fieldSeq->get_Result());

        // Display the next number in this sequence with another SEQ field.
        builder->Write(u", #");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        fieldSeq->Update();

        ASSERT_EQ(u"101", fieldSeq->get_Result());

        // Insert a level 1 heading.
        builder->InsertBreak(BreakType::ParagraphBreak);
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
        builder->Writeln(u"This level 1 heading will reset MySequence to 1");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Normal"));

        // Insert another SEQ field from the same sequence and configure it to reset the count at every heading with 1.
        builder->Write(u"\n#");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        fieldSeq->set_ResetHeadingLevel(u"1");
        fieldSeq->Update();

        // The above heading is a level 1 heading, so the count for this sequence is reset to 1.
        ASSERT_EQ(u" SEQ  MySequence \\s 1", fieldSeq->GetFieldCode());
        ASSERT_EQ(u"1", fieldSeq->get_Result());

        // Move to the next number of this sequence.
        builder->Write(u", #");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        fieldSeq->set_InsertNextNumber(true);
        fieldSeq->Update();

        ASSERT_EQ(u" SEQ  MySequence \\n", fieldSeq->GetFieldCode());
        ASSERT_EQ(u"2", fieldSeq->get_Result());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.SEQ.ResetNumbering.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.SEQ.ResetNumbering.docx");

        ASSERT_EQ(4, doc->get_Range()->get_Fields()->get_Count());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence \\r 100", u"100", fieldSeq);
        ASSERT_EQ(u"MySequence", fieldSeq->get_SequenceIdentifier());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"101", fieldSeq);
        ASSERT_EQ(u"MySequence", fieldSeq->get_SequenceIdentifier());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence \\s 1", u"1", fieldSeq);
        ASSERT_EQ(u"MySequence", fieldSeq->get_SequenceIdentifier());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence \\n", u"2", fieldSeq);
        ASSERT_EQ(u"MySequence", fieldSeq->get_SequenceIdentifier());
    }

    void TocSeqBookmark()
    {
        //ExStart
        //ExFor:FieldSeq
        //ExFor:FieldSeq.BookmarkName
        //ExSummary:Shows how to combine table of contents and sequence fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // A TOC field can create an entry in its table of contents for each SEQ field found in the document.
        // Each entry contains the paragraph that contains the SEQ field,
        // and the number of the page that the field appears on.
        auto fieldToc = System::DynamicCast<FieldToc>(builder->InsertField(FieldType::FieldTOC, true));

        // Configure this TOC field to have a SequenceIdentifier property with a value of "MySequence".
        fieldToc->set_TableOfFiguresLabel(u"MySequence");

        // Configure this TOC field to only pick up SEQ fields that are within the bounds of a bookmark
        // named "TOCBookmark".
        fieldToc->set_BookmarkName(u"TOCBookmark");
        builder->InsertBreak(BreakType::PageBreak);

        ASSERT_EQ(u" TOC  \\c MySequence \\b TOCBookmark", fieldToc->GetFieldCode());

        // SEQ fields display a count that increments at each SEQ field.
        // These fields also maintain separate counts for each unique named sequence
        // identified by the SEQ field's "SequenceIdentifier" property.
        // Insert a SEQ field that has a sequence identifier that matches the TOC's
        // TableOfFiguresLabel property. This field will not create an entry in the TOC since it is outside
        // the bookmark's bounds designated by "BookmarkName".
        builder->Write(u"MySequence #");
        auto fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        builder->Writeln(u", will not show up in the TOC because it is outside of the bookmark.");

        builder->StartBookmark(u"TOCBookmark");

        // This SEQ field's sequence matches the TOC's "TableOfFiguresLabel" property and is within the bookmark's bounds.
        // The paragraph that contains this field will show up in the TOC as an entry.
        builder->Write(u"MySequence #");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        builder->Writeln(u", will show up in the TOC next to the entry for the above caption.");

        // This SEQ field's sequence does not match the TOC's "TableOfFiguresLabel" property,
        // and is within the bounds of the bookmark. Its paragraph will not show up in the TOC as an entry.
        builder->Write(u"MySequence #");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"OtherSequence");
        builder->Writeln(u", will not show up in the TOC because it's from a different sequence identifier.");

        // This SEQ field's sequence matches the TOC's "TableOfFiguresLabel" property and is within the bounds of the bookmark.
        // This field also references another bookmark. The contents of that bookmark will appear in the TOC entry for this SEQ field.
        // The SEQ field itself will not display the contents of that bookmark.
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        fieldSeq->set_BookmarkName(u"SEQBookmark");
        ASSERT_EQ(u" SEQ  MySequence SEQBookmark", fieldSeq->GetFieldCode());

        // Create a bookmark with contents that will show up in the TOC entry due to the above SEQ field referencing it.
        builder->InsertBreak(BreakType::PageBreak);
        builder->StartBookmark(u"SEQBookmark");
        builder->Write(u"MySequence #");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        builder->Writeln(u", text from inside SEQBookmark.");
        builder->EndBookmark(u"SEQBookmark");

        builder->EndBookmark(u"TOCBookmark");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.SEQ.Bookmark.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.SEQ.Bookmark.docx");

        ASSERT_EQ(8, doc->get_Range()->get_Fields()->get_Count());

        fieldToc = System::DynamicCast<FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));
        ArrayPtr<String> pageRefIds = fieldToc->get_Result()
                                          .Split(MakeArray<char16_t>({u' '}))
                                          ->LINQ_Where(static_cast<System::Func<String, bool>>(
                                              static_cast<std::function<bool(String s)>>([](String s) -> bool { return s.StartsWith(u"_Toc"); })))
                                          ->LINQ_ToArray();

        ASSERT_EQ(FieldType::FieldTOC, fieldToc->get_Type());
        ASSERT_EQ(u"MySequence", fieldToc->get_TableOfFiguresLabel());
        TestUtil::VerifyField(FieldType::FieldTOC, u" TOC  \\c MySequence \\b TOCBookmark",
                              String::Format((u"MySequence #2, will show up in the TOC next to the entry for the above caption.\t\u0013 PAGEREF {0} \\h \u0014"
                                              u"2\u0015\r"),
                                             pageRefIds[0]) +
                                  String::Format((u"3MySequence #3, text from inside SEQBookmark.\t\u0013 PAGEREF {0} \\h \u0014"
                                                  u"2\u0015\r"),
                                                 pageRefIds[1]),
                              fieldToc);

        auto fieldPageRef = System::DynamicCast<FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldPageRef, String::Format(u" PAGEREF {0} \\h ", pageRefIds[0]), u"2", fieldPageRef);
        ASSERT_EQ(pageRefIds[0], fieldPageRef->get_BookmarkName());

        fieldPageRef = System::DynamicCast<FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldPageRef, String::Format(u" PAGEREF {0} \\h ", pageRefIds[1]), u"2", fieldPageRef);
        ASSERT_EQ(pageRefIds[1], fieldPageRef->get_BookmarkName());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"1", fieldSeq);
        ASSERT_EQ(u"MySequence", fieldSeq->get_SequenceIdentifier());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(4));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"2", fieldSeq);
        ASSERT_EQ(u"MySequence", fieldSeq->get_SequenceIdentifier());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(5));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  OtherSequence", u"1", fieldSeq);
        ASSERT_EQ(u"OtherSequence", fieldSeq->get_SequenceIdentifier());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(6));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence SEQBookmark", u"3", fieldSeq);
        ASSERT_EQ(u"MySequence", fieldSeq->get_SequenceIdentifier());
        ASSERT_EQ(u"SEQBookmark", fieldSeq->get_BookmarkName());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(7));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"3", fieldSeq);
        ASSERT_EQ(u"MySequence", fieldSeq->get_SequenceIdentifier());
    }

    void FieldCitation_()
    {
        //ExStart
        //ExFor:FieldCitation
        //ExFor:FieldCitation.AnotherSourceTag
        //ExFor:FieldCitation.FormatLanguageId
        //ExFor:FieldCitation.PageNumber
        //ExFor:FieldCitation.Prefix
        //ExFor:FieldCitation.SourceTag
        //ExFor:FieldCitation.Suffix
        //ExFor:FieldCitation.SuppressAuthor
        //ExFor:FieldCitation.SuppressTitle
        //ExFor:FieldCitation.SuppressYear
        //ExFor:FieldCitation.VolumeNumber
        //ExFor:FieldBibliography
        //ExFor:FieldBibliography.FormatLanguageId
        //ExSummary:Shows how to work with CITATION and BIBLIOGRAPHY fields.
        // Open a document containing bibliographical sources that we can find in
        // Microsoft Word via References -> Citations & Bibliography -> Manage Sources.
        auto doc = MakeObject<Document>(MyDir + u"Bibliography.docx");
        ASSERT_EQ(2, doc->get_Range()->get_Fields()->get_Count());
        //ExSkip

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Text to be cited with one source.");

        // Create a citation with just the page number and the author of the referenced book.
        auto fieldCitation = System::DynamicCast<FieldCitation>(builder->InsertField(FieldType::FieldCitation, true));

        // We refer to sources using their tag names.
        fieldCitation->set_SourceTag(u"Book1");
        fieldCitation->set_PageNumber(u"85");
        fieldCitation->set_SuppressAuthor(false);
        fieldCitation->set_SuppressTitle(true);
        fieldCitation->set_SuppressYear(true);

        ASSERT_EQ(u" CITATION  Book1 \\p 85 \\t \\y", fieldCitation->GetFieldCode());

        // Create a more detailed citation which cites two sources.
        builder->InsertParagraph();
        builder->Write(u"Text to be cited with two sources.");
        fieldCitation = System::DynamicCast<FieldCitation>(builder->InsertField(FieldType::FieldCitation, true));
        fieldCitation->set_SourceTag(u"Book1");
        fieldCitation->set_AnotherSourceTag(u"Book2");
        fieldCitation->set_FormatLanguageId(u"en-US");
        fieldCitation->set_PageNumber(u"19");
        fieldCitation->set_Prefix(u"Prefix ");
        fieldCitation->set_Suffix(u" Suffix");
        fieldCitation->set_SuppressAuthor(false);
        fieldCitation->set_SuppressTitle(false);
        fieldCitation->set_SuppressYear(false);
        fieldCitation->set_VolumeNumber(u"VII");

        ASSERT_EQ(u" CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII", fieldCitation->GetFieldCode());

        // We can use a BIBLIOGRAPHY field to display all the sources within the document.
        builder->InsertBreak(BreakType::PageBreak);
        auto fieldBibliography = System::DynamicCast<FieldBibliography>(builder->InsertField(FieldType::FieldBibliography, true));
        fieldBibliography->set_FormatLanguageId(u"1124");

        ASSERT_EQ(u" BIBLIOGRAPHY  \\l 1124", fieldBibliography->GetFieldCode());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.CITATION.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.CITATION.docx");

        ASSERT_EQ(5, doc->get_Range()->get_Fields()->get_Count());

        fieldCitation = System::DynamicCast<FieldCitation>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldCitation, u" CITATION  Book1 \\p 85 \\t \\y", u" (Doe, p. 85)", fieldCitation);
        ASSERT_EQ(u"Book1", fieldCitation->get_SourceTag());
        ASSERT_EQ(u"85", fieldCitation->get_PageNumber());
        ASSERT_FALSE(fieldCitation->get_SuppressAuthor());
        ASSERT_TRUE(fieldCitation->get_SuppressTitle());
        ASSERT_TRUE(fieldCitation->get_SuppressYear());

        fieldCitation = System::DynamicCast<FieldCitation>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldCitation, u" CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII",
                              u" (Doe, 2018; Prefix Cardholder, 2018, VII:19 Suffix)", fieldCitation);
        ASSERT_EQ(u"Book1", fieldCitation->get_SourceTag());
        ASSERT_EQ(u"Book2", fieldCitation->get_AnotherSourceTag());
        ASSERT_EQ(u"en-US", fieldCitation->get_FormatLanguageId());
        ASSERT_EQ(u"Prefix ", fieldCitation->get_Prefix());
        ASSERT_EQ(u" Suffix", fieldCitation->get_Suffix());
        ASSERT_EQ(u"19", fieldCitation->get_PageNumber());
        ASSERT_FALSE(fieldCitation->get_SuppressAuthor());
        ASSERT_FALSE(fieldCitation->get_SuppressTitle());
        ASSERT_FALSE(fieldCitation->get_SuppressYear());
        ASSERT_EQ(u"VII", fieldCitation->get_VolumeNumber());

        fieldBibliography = System::DynamicCast<FieldBibliography>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldBibliography, u" BIBLIOGRAPHY  \\l 1124",
                              u"Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r",
                              fieldBibliography);
        ASSERT_EQ(u"1124", fieldBibliography->get_FormatLanguageId());

        fieldCitation = System::DynamicCast<FieldCitation>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldCitation, u" CITATION Book1 \\l 1033 ", u"(Doe, 2018)", fieldCitation);
        ASSERT_EQ(u"Book1", fieldCitation->get_SourceTag());
        ASSERT_EQ(u"1033", fieldCitation->get_FormatLanguageId());

        fieldBibliography = System::DynamicCast<FieldBibliography>(doc->get_Range()->get_Fields()->idx_get(4));

        TestUtil::VerifyField(FieldType::FieldBibliography, u" BIBLIOGRAPHY ",
                              u"Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r",
                              fieldBibliography);
    }

    void FieldData_()
    {
        //ExStart
        //ExFor:FieldData
        //ExSummary:Shows how to insert a DATA field into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto field = System::DynamicCast<FieldData>(builder->InsertField(FieldType::FieldData, true));
        ASSERT_EQ(u" DATA ", field->GetFieldCode());
        //ExEnd

        TestUtil::VerifyField(FieldType::FieldData, u" DATA ", String::Empty, DocumentHelper::SaveOpen(doc)->get_Range()->get_Fields()->idx_get(0));
    }

    void FieldInclude_()
    {
        //ExStart
        //ExFor:FieldInclude
        //ExFor:FieldInclude.BookmarkName
        //ExFor:FieldInclude.LockFields
        //ExFor:FieldInclude.SourceFullName
        //ExFor:FieldInclude.TextConverter
        //ExSummary:Shows how to create an INCLUDE field, and set its properties.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // We can use an INCLUDE field to import a portion of another document in the local file system.
        // The bookmark from the other document that we reference with this field contains this imported portion.
        auto field = System::DynamicCast<FieldInclude>(builder->InsertField(FieldType::FieldInclude, true));
        field->set_SourceFullName(MyDir + u"Bookmarks.docx");
        field->set_BookmarkName(u"MyBookmark1");
        field->set_LockFields(false);
        field->set_TextConverter(u"Microsoft Word");

        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(field->GetFieldCode(), u" INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\"")->get_Success());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.INCLUDE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.INCLUDE.docx");
        field = System::DynamicCast<FieldInclude>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(FieldType::FieldInclude, field->get_Type());
        ASSERT_EQ(u"First bookmark.", field->get_Result());
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(field->GetFieldCode(), u" INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\"")->get_Success());

        ASSERT_EQ(MyDir + u"Bookmarks.docx", field->get_SourceFullName());
        ASSERT_EQ(u"MyBookmark1", field->get_BookmarkName());
        ASSERT_FALSE(field->get_LockFields());
        ASSERT_EQ(u"Microsoft Word", field->get_TextConverter());
    }

    void FieldIncludePicture_()
    {
        //ExStart
        //ExFor:FieldIncludePicture
        //ExFor:FieldIncludePicture.GraphicFilter
        //ExFor:FieldIncludePicture.IsLinked
        //ExFor:FieldIncludePicture.ResizeHorizontally
        //ExFor:FieldIncludePicture.ResizeVertically
        //ExFor:FieldIncludePicture.SourceFullName
        //ExFor:FieldImport
        //ExFor:FieldImport.GraphicFilter
        //ExFor:FieldImport.IsLinked
        //ExFor:FieldImport.SourceFullName
        //ExSummary:Shows how to insert images using IMPORT and INCLUDEPICTURE fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are two similar field types that we can use to display images linked from the local file system.
        // 1 -  The INCLUDEPICTURE field:
        auto fieldIncludePicture = System::DynamicCast<FieldIncludePicture>(builder->InsertField(FieldType::FieldIncludePicture, true));
        fieldIncludePicture->set_SourceFullName(ImageDir + u"Transparent background logo.png");

        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(fieldIncludePicture->GetFieldCode(), u" INCLUDEPICTURE  .*")->get_Success());

        // Apply the PNG32.FLT filter.
        fieldIncludePicture->set_GraphicFilter(u"PNG32");
        fieldIncludePicture->set_IsLinked(true);
        fieldIncludePicture->set_ResizeHorizontally(true);
        fieldIncludePicture->set_ResizeVertically(true);

        // 2 -  The IMPORT field:
        auto fieldImport = System::DynamicCast<FieldImport>(builder->InsertField(FieldType::FieldImport, true));
        fieldImport->set_SourceFullName(ImageDir + u"Transparent background logo.png");
        fieldImport->set_GraphicFilter(u"PNG32");
        fieldImport->set_IsLinked(true);

        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(fieldImport->GetFieldCode(), u" IMPORT  .* \\\\c PNG32 \\\\d")->get_Success());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.IMPORT.INCLUDEPICTURE.docx");
        //ExEnd

        ASSERT_EQ(ImageDir + u"Transparent background logo.png", fieldIncludePicture->get_SourceFullName());
        ASSERT_EQ(u"PNG32", fieldIncludePicture->get_GraphicFilter());
        ASSERT_TRUE(fieldIncludePicture->get_IsLinked());
        ASSERT_TRUE(fieldIncludePicture->get_ResizeHorizontally());
        ASSERT_TRUE(fieldIncludePicture->get_ResizeVertically());

        ASSERT_EQ(ImageDir + u"Transparent background logo.png", fieldImport->get_SourceFullName());
        ASSERT_EQ(u"PNG32", fieldImport->get_GraphicFilter());
        ASSERT_TRUE(fieldImport->get_IsLinked());

        doc = MakeObject<Document>(ArtifactsDir + u"Field.IMPORT.INCLUDEPICTURE.docx");

        // The INCLUDEPICTURE fields have been converted into shapes with linked images during loading.
        ASSERT_EQ(0, doc->get_Range()->get_Fields()->get_Count());
        ASSERT_EQ(2, doc->GetChildNodes(NodeType::Shape, true)->get_Count());

        auto image = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_TRUE(image->get_IsImage());
        ASSERT_TRUE(image->get_ImageData()->get_ImageBytes() == nullptr);
        ASSERT_EQ(ImageDir + u"Transparent background logo.png", image->get_ImageData()->get_SourceFullName().Replace(u"%20", u" "));

        image = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        ASSERT_TRUE(image->get_IsImage());
        ASSERT_TRUE(image->get_ImageData()->get_ImageBytes() == nullptr);
        ASSERT_EQ(ImageDir + u"Transparent background logo.png", image->get_ImageData()->get_SourceFullName().Replace(u"%20", u" "));
    }

    //ExStart
    //ExFor:FieldIncludeText
    //ExFor:FieldIncludeText.BookmarkName
    //ExFor:FieldIncludeText.Encoding
    //ExFor:FieldIncludeText.LockFields
    //ExFor:FieldIncludeText.MimeType
    //ExFor:FieldIncludeText.NamespaceMappings
    //ExFor:FieldIncludeText.SourceFullName
    //ExFor:FieldIncludeText.TextConverter
    //ExFor:FieldIncludeText.XPath
    //ExFor:FieldIncludeText.XslTransformation
    //ExSummary:Shows how to create an INCLUDETEXT field, and set its properties.
    void FieldIncludeText_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are two ways to use INCLUDETEXT fields to display the contents of an XML file in the local file system.
        // 1 -  Perform an XSL transformation on an XML document:
        SharedPtr<FieldIncludeText> fieldIncludeText =
            CreateFieldIncludeText(builder, MyDir + u"CD collection data.xml", false, u"text/xml", u"XML", u"ISO-8859-1");
        fieldIncludeText->set_XslTransformation(MyDir + u"CD collection XSL transformation.xsl");

        builder->Writeln();

        // 2 -  Use an XPath to take specific elements from an XML document:
        fieldIncludeText = CreateFieldIncludeText(builder, MyDir + u"CD collection data.xml", false, u"text/xml", u"XML", u"ISO-8859-1");
        fieldIncludeText->set_NamespaceMappings(u"xmlns:n='myNamespace'");
        fieldIncludeText->set_XPath(u"/catalog/cd/title");

        doc->Save(ArtifactsDir + u"Field.INCLUDETEXT.docx");
        TestFieldIncludeText(MakeObject<Document>(ArtifactsDir + u"Field.INCLUDETEXT.docx"));
        //ExSkip
    }

    /// <summary>
    /// Use a document builder to insert an INCLUDETEXT field with custom properties.
    /// </summary>
    SharedPtr<FieldIncludeText> CreateFieldIncludeText(SharedPtr<DocumentBuilder> builder, String sourceFullName, bool lockFields, String mimeType,
                                                       String textConverter, String encoding)
    {
        auto fieldIncludeText = System::DynamicCast<FieldIncludeText>(builder->InsertField(FieldType::FieldIncludeText, true));
        fieldIncludeText->set_SourceFullName(sourceFullName);
        fieldIncludeText->set_LockFields(lockFields);
        fieldIncludeText->set_MimeType(mimeType);
        fieldIncludeText->set_TextConverter(textConverter);
        fieldIncludeText->set_Encoding(encoding);

        return fieldIncludeText;
    }
    //ExEnd

    void TestFieldIncludeText(SharedPtr<Document> doc)
    {
        doc = DocumentHelper::SaveOpen(doc);

        auto fieldIncludeText = System::DynamicCast<FieldIncludeText>(doc->get_Range()->get_Fields()->idx_get(0));
        ASSERT_EQ(MyDir + u"CD collection data.xml", fieldIncludeText->get_SourceFullName());
        ASSERT_EQ(MyDir + u"CD collection XSL transformation.xsl", fieldIncludeText->get_XslTransformation());
        ASSERT_FALSE(fieldIncludeText->get_LockFields());
        ASSERT_EQ(u"text/xml", fieldIncludeText->get_MimeType());
        ASSERT_EQ(u"XML", fieldIncludeText->get_TextConverter());
        ASSERT_EQ(u"ISO-8859-1", fieldIncludeText->get_Encoding());
        ASSERT_EQ(String(u" INCLUDETEXT  \"") + MyDir.Replace(u"\\", u"\\\\") + u"CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\t \"" +
                      MyDir.Replace(u"\\", u"\\\\") + u"CD collection XSL transformation.xsl\"",
                  fieldIncludeText->GetFieldCode());
        ASSERT_TRUE(fieldIncludeText->get_Result().StartsWith(u"My CD Collection"));

        auto cdCollectionData = MakeObject<System::Xml::XmlDocument>();
        cdCollectionData->LoadXml(System::IO::File::ReadAllText(MyDir + u"CD collection data.xml"));
        SharedPtr<System::Xml::XmlNode> catalogData = cdCollectionData->get_ChildNodes()->idx_get(0);

        auto cdCollectionXslTransformation = MakeObject<System::Xml::XmlDocument>();
        cdCollectionXslTransformation->LoadXml(System::IO::File::ReadAllText(MyDir + u"CD collection XSL transformation.xsl"));

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        auto manager = MakeObject<System::Xml::XmlNamespaceManager>(cdCollectionXslTransformation->get_NameTable());
        manager->AddNamespace(u"xsl", u"http://www.w3.org/1999/XSL/Transform");

        for (int i = 0; i < table->get_Rows()->get_Count(); i++)
        {
            for (int j = 0; j < table->get_Rows()->idx_get(i)->get_Count(); j++)
            {
                if (i == 0)
                {
                    // When on the first row from the input document's table, ensure that all table's cells match all XML element Names.
                    for (int k = 0; k < table->get_Rows()->get_Count() - 1; k++)
                    {
                        ASSERT_EQ(catalogData->get_ChildNodes()->idx_get(k)->get_ChildNodes()->idx_get(j)->get_Name(),
                                  table->get_Rows()->idx_get(i)->get_Cells()->idx_get(j)->GetText().Replace(ControlChar::Cell(), String::Empty).ToLower());
                    }

                    // Also, make sure that the whole first row has the same color as the XSL transform.
                    ASSERT_EQ(cdCollectionXslTransformation->SelectNodes(u"//xsl:stylesheet/xsl:template/html/body/table/tr", manager)
                                  ->idx_get(0)
                                  ->get_Attributes()
                                  ->GetNamedItem(u"bgcolor")
                                  ->get_Value(),
                              System::Drawing::ColorTranslator::ToHtml(
                                  table->get_Rows()->idx_get(i)->get_Cells()->idx_get(j)->get_CellFormat()->get_Shading()->get_BackgroundPatternColor())
                                  .ToLower());
                }
                else
                {
                    // When on all other rows of the input document's table, ensure that cell contents match XML element Values.
                    ASSERT_EQ(catalogData->get_ChildNodes()->idx_get(i - 1)->get_ChildNodes()->idx_get(j)->get_FirstChild()->get_Value(),
                              table->get_Rows()->idx_get(i)->get_Cells()->idx_get(j)->GetText().Replace(ControlChar::Cell(), String::Empty));
                    ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty,
                                     table->get_Rows()->idx_get(i)->get_Cells()->idx_get(j)->get_CellFormat()->get_Shading()->get_BackgroundPatternColor());
                }

                ASPOSE_ASSERT_EQ(System::Double::Parse(cdCollectionXslTransformation->SelectNodes(u"//xsl:stylesheet/xsl:template/html/body/table", manager)
                                                           ->idx_get(0)
                                                           ->get_Attributes()
                                                           ->GetNamedItem(u"border")
                                                           ->get_Value()) *
                                     0.75,
                                 table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Bottom()->get_LineWidth());
            }
        }

        fieldIncludeText = System::DynamicCast<FieldIncludeText>(doc->get_Range()->get_Fields()->idx_get(1));
        ASSERT_EQ(MyDir + u"CD collection data.xml", fieldIncludeText->get_SourceFullName());
        ASSERT_TRUE(fieldIncludeText->get_XslTransformation() == nullptr);
        ASSERT_FALSE(fieldIncludeText->get_LockFields());
        ASSERT_EQ(u"text/xml", fieldIncludeText->get_MimeType());
        ASSERT_EQ(u"XML", fieldIncludeText->get_TextConverter());
        ASSERT_EQ(u"ISO-8859-1", fieldIncludeText->get_Encoding());
        ASSERT_EQ(String(u" INCLUDETEXT  \"") + MyDir.Replace(u"\\", u"\\\\") +
                      u"CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\n xmlns:n='myNamespace' \\x /catalog/cd/title",
                  fieldIncludeText->GetFieldCode());

        String expectedFieldResult = u"";
        for (int i = 0; i < catalogData->get_ChildNodes()->get_Count(); i++)
        {
            expectedFieldResult += catalogData->get_ChildNodes()->idx_get(i)->get_ChildNodes()->idx_get(0)->get_ChildNodes()->idx_get(0)->get_Value();
        }

        ASSERT_EQ(expectedFieldResult, fieldIncludeText->get_Result());
    }

    void FieldHyperlink_()
    {
        //ExStart
        //ExFor:FieldHyperlink
        //ExFor:FieldHyperlink.Address
        //ExFor:FieldHyperlink.IsImageMap
        //ExFor:FieldHyperlink.OpenInNewWindow
        //ExFor:FieldHyperlink.ScreenTip
        //ExFor:FieldHyperlink.SubAddress
        //ExFor:FieldHyperlink.Target
        //ExSummary:Shows how to use HYPERLINK fields to link to documents in the local file system.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto field = System::DynamicCast<FieldHyperlink>(builder->InsertField(FieldType::FieldHyperlink, true));

        // When we click this HYPERLINK field in Microsoft Word,
        // it will open the linked document and then place the cursor at the specified bookmark.
        field->set_Address(MyDir + u"Bookmarks.docx");
        field->set_SubAddress(u"MyBookmark3");
        field->set_ScreenTip(String(u"Open ") + field->get_Address() + u" on bookmark " + field->get_SubAddress() + u" in a new window");

        builder->Writeln();

        // When we click this HYPERLINK field in Microsoft Word,
        // it will open the linked document, and automatically scroll down to the specified iframe.
        field = System::DynamicCast<FieldHyperlink>(builder->InsertField(FieldType::FieldHyperlink, true));
        field->set_Address(MyDir + u"Iframes.html");
        field->set_ScreenTip(String(u"Open ") + field->get_Address());
        field->set_Target(u"iframe_3");
        field->set_OpenInNewWindow(true);
        field->set_IsImageMap(false);

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.HYPERLINK.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.HYPERLINK.docx");
        field = System::DynamicCast<FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldHyperlink,
                              String(u" HYPERLINK \"") + MyDir.Replace(u"\\", u"\\\\") + u"Bookmarks.docx\" \\l \"MyBookmark3\" \\o \"Open " + MyDir +
                                  u"Bookmarks.docx on bookmark MyBookmark3 in a new window\" ",
                              MyDir + u"Bookmarks.docx - MyBookmark3", field);
        ASSERT_EQ(MyDir + u"Bookmarks.docx", field->get_Address());
        ASSERT_EQ(u"MyBookmark3", field->get_SubAddress());
        ASSERT_EQ(String(u"Open ") + field->get_Address().Replace(u"\\", String::Empty) + u" on bookmark " + field->get_SubAddress() + u" in a new window",
                  field->get_ScreenTip());

        field = System::DynamicCast<FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldHyperlink,
                              String(u" HYPERLINK \"file:///") + MyDir.Replace(u"\\", u"\\\\").Replace(u" ", u"%20") +
                                  u"Iframes.html\" \\t \"iframe_3\" \\o \"Open " + MyDir.Replace(u"\\", u"\\\\") + u"Iframes.html\" ",
                              MyDir + u"Iframes.html", field);
        ASSERT_EQ(String(u"file:///") + MyDir.Replace(u" ", u"%20") + u"Iframes.html", field->get_Address());
        ASSERT_EQ(String(u"Open ") + MyDir + u"Iframes.html", field->get_ScreenTip());
        ASSERT_EQ(u"iframe_3", field->get_Target());
        ASSERT_FALSE(field->get_OpenInNewWindow());
        ASSERT_FALSE(field->get_IsImageMap());
    }

    void FieldIndexFilter()
    {
        //ExStart
        //ExFor:FieldIndex
        //ExFor:FieldIndex.BookmarkName
        //ExFor:FieldIndex.EntryType
        //ExFor:FieldXE
        //ExFor:FieldXE.EntryType
        //ExFor:FieldXE.Text
        //ExSummary:Shows how to create an INDEX field, and then use XE fields to populate it with entries.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side
        // and the page containing the XE field on the right.
        // If the XE fields have the same value in their "Text" property,
        // the INDEX field will group them into one entry.
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        // Configure the INDEX field only to display XE fields that are within the bounds
        // of a bookmark named "MainBookmark", and whose "EntryType" properties have a value of "A".
        // For both INDEX and XE fields, the "EntryType" property only uses the first character of its string value.
        index->set_BookmarkName(u"MainBookmark");
        index->set_EntryType(u"A");

        ASSERT_EQ(u" INDEX  \\b MainBookmark \\f A", index->GetFieldCode());

        // On a new page, start the bookmark with a name that matches the value
        // of the INDEX field's "BookmarkName" property.
        builder->InsertBreak(BreakType::PageBreak);
        builder->StartBookmark(u"MainBookmark");

        // The INDEX field will pick up this entry because it is inside the bookmark,
        // and its entry type also matches the INDEX field's entry type.
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Index entry 1");
        indexEntry->set_EntryType(u"A");

        ASSERT_EQ(u" XE  \"Index entry 1\" \\f A", indexEntry->GetFieldCode());

        // Insert an XE field that will not appear in the INDEX because the entry types do not match.
        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Index entry 2");
        indexEntry->set_EntryType(u"B");

        // End the bookmark and insert an XE field afterwards.
        // It is of the same type as the INDEX field, but will not appear
        // since it is outside the bookmark's boundaries.
        builder->EndBookmark(u"MainBookmark");
        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Index entry 3");
        indexEntry->set_EntryType(u"A");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.INDEX.XE.Filtering.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.INDEX.XE.Filtering.docx");
        index = System::DynamicCast<FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldIndex, u" INDEX  \\b MainBookmark \\f A", u"Index entry 1, 2\r", index);
        ASSERT_EQ(u"MainBookmark", index->get_BookmarkName());
        ASSERT_EQ(u"A", index->get_EntryType());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  \"Index entry 1\" \\f A", String::Empty, indexEntry);
        ASSERT_EQ(u"Index entry 1", indexEntry->get_Text());
        ASSERT_EQ(u"A", indexEntry->get_EntryType());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  \"Index entry 2\" \\f B", String::Empty, indexEntry);
        ASSERT_EQ(u"Index entry 2", indexEntry->get_Text());
        ASSERT_EQ(u"B", indexEntry->get_EntryType());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  \"Index entry 3\" \\f A", String::Empty, indexEntry);
        ASSERT_EQ(u"Index entry 3", indexEntry->get_Text());
        ASSERT_EQ(u"A", indexEntry->get_EntryType());
    }

    void FieldIndexFormatting()
    {
        //ExStart
        //ExFor:FieldIndex
        //ExFor:FieldIndex.Heading
        //ExFor:FieldIndex.NumberOfColumns
        //ExFor:FieldIndex.LanguageId
        //ExFor:FieldIndex.LetterRange
        //ExFor:FieldXE
        //ExFor:FieldXE.IsBold
        //ExFor:FieldXE.IsItalic
        //ExFor:FieldXE.Text
        //ExSummary:Shows how to populate an INDEX field with entries using XE fields, and also modify its appearance.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // If the XE fields have the same value in their "Text" property,
        // the INDEX field will group them into one entry.
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));
        index->set_LanguageId(u"1033");

        // Setting this property's value to "A" will group all the entries by their first letter,
        // and place that letter in uppercase above each group.
        index->set_Heading(u"A");

        // Set the table created by the INDEX field to span over 2 columns.
        index->set_NumberOfColumns(u"2");

        // Set any entries with starting letters outside the "a-c" character range to be omitted.
        index->set_LetterRange(u"a-c");

        ASSERT_EQ(u" INDEX  \\z 1033 \\h A \\c 2 \\p a-c", index->GetFieldCode());

        // These next two XE fields will show up under the "A" heading,
        // with their respective text stylings also applied to their page numbers.
        builder->InsertBreak(BreakType::PageBreak);
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Apple");
        indexEntry->set_IsItalic(true);

        ASSERT_EQ(u" XE  Apple \\i", indexEntry->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Apricot");
        indexEntry->set_IsBold(true);

        ASSERT_EQ(u" XE  Apricot \\b", indexEntry->GetFieldCode());

        // Both the next two XE fields will be under a "B" and "C" heading in the INDEX fields table of contents.
        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Banana");

        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Cherry");

        // INDEX fields sort all entries alphabetically, so this entry will show up under "A" with the other two.
        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Avocado");

        // This entry will not appear because it starts with the letter "D",
        // which is outside the "a-c" character range that the INDEX field's LetterRange property defines.
        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Durian");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.INDEX.XE.Formatting.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.INDEX.XE.Formatting.docx");
        index = System::DynamicCast<FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(u"1033", index->get_LanguageId());
        ASSERT_EQ(u"A", index->get_Heading());
        ASSERT_EQ(u"2", index->get_NumberOfColumns());
        ASSERT_EQ(u"a-c", index->get_LetterRange());
        ASSERT_EQ(u" INDEX  \\z 1033 \\h A \\c 2 \\p a-c", index->GetFieldCode());
        ASSERT_EQ(String(u"\fA\r") + u"Apple, 2\r" + u"Apricot, 3\r" + u"Avocado, 6\r" + u"B\r" + u"Banana, 4\r" + u"C\r" + u"Cherry, 5\r\f",
                  index->get_Result());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Apple \\i", String::Empty, indexEntry);
        ASSERT_EQ(u"Apple", indexEntry->get_Text());
        ASSERT_FALSE(indexEntry->get_IsBold());
        ASSERT_TRUE(indexEntry->get_IsItalic());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Apricot \\b", String::Empty, indexEntry);
        ASSERT_EQ(u"Apricot", indexEntry->get_Text());
        ASSERT_TRUE(indexEntry->get_IsBold());
        ASSERT_FALSE(indexEntry->get_IsItalic());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Banana", String::Empty, indexEntry);
        ASSERT_EQ(u"Banana", indexEntry->get_Text());
        ASSERT_FALSE(indexEntry->get_IsBold());
        ASSERT_FALSE(indexEntry->get_IsItalic());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(4));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Cherry", String::Empty, indexEntry);
        ASSERT_EQ(u"Cherry", indexEntry->get_Text());
        ASSERT_FALSE(indexEntry->get_IsBold());
        ASSERT_FALSE(indexEntry->get_IsItalic());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(5));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Avocado", String::Empty, indexEntry);
        ASSERT_EQ(u"Avocado", indexEntry->get_Text());
        ASSERT_FALSE(indexEntry->get_IsBold());
        ASSERT_FALSE(indexEntry->get_IsItalic());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(6));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Durian", String::Empty, indexEntry);
        ASSERT_EQ(u"Durian", indexEntry->get_Text());
        ASSERT_FALSE(indexEntry->get_IsBold());
        ASSERT_FALSE(indexEntry->get_IsItalic());
    }

    void FieldIndexSequence()
    {
        //ExStart
        //ExFor:FieldIndex.HasSequenceName
        //ExFor:FieldIndex.SequenceName
        //ExFor:FieldIndex.SequenceSeparator
        //ExSummary:Shows how to split a document into portions by combining INDEX and SEQ fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // If the XE fields have the same value in their "Text" property,
        // the INDEX field will group them into one entry.
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        // In the SequenceName property, name a SEQ field sequence. Each entry of this INDEX field will now also display
        // the number that the sequence count is on at the XE field location that created this entry.
        index->set_SequenceName(u"MySequence");

        // Set text that will around the sequence and page numbers to explain their meaning to the user.
        // An entry created with this configuration will display something like "MySequence at 1 on page 1" at its page number.
        // PageNumberSeparator and SequenceSeparator cannot be longer than 15 characters.
        index->set_PageNumberSeparator(u"\tMySequence at ");
        index->set_SequenceSeparator(u" on page ");
        ASSERT_TRUE(index->get_HasSequenceName());

        ASSERT_EQ(u" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index->GetFieldCode());

        // SEQ fields display a count that increments at each SEQ field.
        // These fields also maintain separate counts for each unique named sequence
        // identified by the SEQ field's "SequenceIdentifier" property.
        // Insert a SEQ field which moves the "MySequence" sequence to 1.
        // This field no different from normal document text. It will not appear on an INDEX field's table of contents.
        builder->InsertBreak(BreakType::PageBreak);
        auto sequenceField = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        sequenceField->set_SequenceIdentifier(u"MySequence");

        ASSERT_EQ(u" SEQ  MySequence", sequenceField->GetFieldCode());

        // Insert an XE field which will create an entry in the INDEX field.
        // Since "MySequence" is at 1 and this XE field is on page 2, along with the custom separators we defined above,
        // this field's INDEX entry will display "Cat" on the left side, and "MySequence at 1 on page 2" on the right.
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Cat");

        ASSERT_EQ(u" XE  Cat", indexEntry->GetFieldCode());

        // Insert a page break and use SEQ fields to advance "MySequence" to 3.
        builder->InsertBreak(BreakType::PageBreak);
        sequenceField = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        sequenceField->set_SequenceIdentifier(u"MySequence");
        sequenceField = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        sequenceField->set_SequenceIdentifier(u"MySequence");

        // Insert an XE field with the same Text property as the one above.
        // The INDEX entry will group XE fields with matching values in the "Text" property
        // into one entry as opposed to making an entry for each XE field.
        // Since we are on page 2 with "MySequence" at 3, ", 3 on page 3" will be appended to the same INDEX entry as above.
        // The page number portion of that INDEX entry will now display "MySequence at 1 on page 2, 3 on page 3".
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Cat");

        // Insert an XE field with a new and unique Text property value.
        // This will add a new entry, with MySequence at 3 on page 4.
        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Dog");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.INDEX.XE.Sequence.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.INDEX.XE.Sequence.docx");
        index = System::DynamicCast<FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(u"MySequence", index->get_SequenceName());
        ASSERT_EQ(u"\tMySequence at ", index->get_PageNumberSeparator());
        ASSERT_EQ(u" on page ", index->get_SequenceSeparator());
        ASSERT_TRUE(index->get_HasSequenceName());
        ASSERT_EQ(u" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index->GetFieldCode());
        ASSERT_EQ(String(u"Cat\tMySequence at 1 on page 2, 3 on page 3\r") + u"Dog\tMySequence at 3 on page 4\r", index->get_Result());

        ASSERT_EQ(
            3, doc->get_Range()
                   ->get_Fields()
                   ->LINQ_Where(static_cast<System::Func<SharedPtr<Field>, bool>>(static_cast<std::function<bool(SharedPtr<Aspose::Words::Fields::Field> f)>>(
                       [](SharedPtr<Aspose::Words::Fields::Field> f) -> bool { return f->get_Type() == FieldType::FieldSequence; })))
                   ->LINQ_Count());
    }

    void FieldIndexPageNumberSeparator()
    {
        //ExStart
        //ExFor:FieldIndex.HasPageNumberSeparator
        //ExFor:FieldIndex.PageNumberSeparator
        //ExFor:FieldIndex.PageNumberListSeparator
        //ExSummary:Shows how to edit the page number separator in an INDEX field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // The INDEX entry will group XE fields with matching values in the "Text" property
        // into one entry as opposed to making an entry for each XE field.
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        // If our INDEX field has an entry for a group of XE fields,
        // this entry will display the number of each page that contains an XE field that belongs to this group.
        // We can set custom separators to customize the appearance of these page numbers.
        index->set_PageNumberSeparator(u", on page(s) ");
        index->set_PageNumberListSeparator(u" & ");

        ASSERT_EQ(u" INDEX  \\e \", on page(s) \" \\l \" & \"", index->GetFieldCode());
        ASSERT_TRUE(index->get_HasPageNumberSeparator());

        // After we insert these XE fields, the INDEX field will display "First entry, on page(s) 2 & 3 & 4".
        builder->InsertBreak(BreakType::PageBreak);
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"First entry");

        ASSERT_EQ(u" XE  \"First entry\"", indexEntry->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"First entry");

        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"First entry");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.INDEX.XE.PageNumberList.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.INDEX.XE.PageNumberList.docx");
        index = System::DynamicCast<FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldIndex, u" INDEX  \\e \", on page(s) \" \\l \" & \"", u"First entry, on page(s) 2 & 3 & 4\r", index);
        ASSERT_EQ(u", on page(s) ", index->get_PageNumberSeparator());
        ASSERT_EQ(u" & ", index->get_PageNumberListSeparator());
        ASSERT_TRUE(index->get_HasPageNumberSeparator());
    }

    void FieldIndexPageRangeBookmark()
    {
        //ExStart
        //ExFor:FieldIndex.PageRangeSeparator
        //ExFor:FieldXE.PageRangeBookmarkName
        //ExSummary:Shows how to specify a bookmark's spanned pages as a page range for an INDEX field entry.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // The INDEX entry will collect all XE fields with matching values in the "Text" property
        // into one entry as opposed to making an entry for each XE field.
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        // For INDEX entries that display page ranges, we can specify a separator string
        // which will appear between the number of the first page, and the number of the last.
        index->set_PageNumberSeparator(u", on page(s) ");
        index->set_PageRangeSeparator(u" to ");

        ASSERT_EQ(u" INDEX  \\e \", on page(s) \" \\g \" to \"", index->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"My entry");

        // If an XE field names a bookmark using the PageRangeBookmarkName property,
        // its INDEX entry will show the range of pages that the bookmark spans
        // instead of the number of the page that contains the XE field.
        indexEntry->set_PageRangeBookmarkName(u"MyBookmark");

        ASSERT_EQ(u" XE  \"My entry\" \\r MyBookmark", indexEntry->GetFieldCode());
        ASSERT_EQ(u"MyBookmark", indexEntry->get_PageRangeBookmarkName());

        // Insert a bookmark that starts on page 3 and ends on page 5.
        // The INDEX entry for the XE field that references this bookmark will display this page range.
        // In our table, the INDEX entry will display "My entry, on page(s) 3 to 5".
        builder->InsertBreak(BreakType::PageBreak);
        builder->StartBookmark(u"MyBookmark");
        builder->Write(u"Start of MyBookmark");
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::PageBreak);
        builder->Write(u"End of MyBookmark");
        builder->EndBookmark(u"MyBookmark");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.INDEX.XE.PageRangeBookmark.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.INDEX.XE.PageRangeBookmark.docx");
        index = System::DynamicCast<FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldIndex, u" INDEX  \\e \", on page(s) \" \\g \" to \"", u"My entry, on page(s) 3 to 5\r", index);
        ASSERT_EQ(u", on page(s) ", index->get_PageNumberSeparator());
        ASSERT_EQ(u" to ", index->get_PageRangeSeparator());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  \"My entry\" \\r MyBookmark", String::Empty, indexEntry);
        ASSERT_EQ(u"My entry", indexEntry->get_Text());
        ASSERT_EQ(u"MyBookmark", indexEntry->get_PageRangeBookmarkName());
    }

    void FieldIndexCrossReferenceSeparator()
    {
        //ExStart
        //ExFor:FieldIndex.CrossReferenceSeparator
        //ExFor:FieldXE.PageNumberReplacement
        //ExSummary:Shows how to define cross references in an INDEX field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // The INDEX entry will collect all XE fields with matching values in the "Text" property
        // into one entry as opposed to making an entry for each XE field.
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        // We can configure an XE field to get its INDEX entry to display a string instead of a page number.
        // First, for entries that substitute a page number with a string,
        // specify a custom separator between the XE field's Text property value and the string.
        index->set_CrossReferenceSeparator(u", see: ");

        ASSERT_EQ(u" INDEX  \\k \", see: \"", index->GetFieldCode());

        // Insert an XE field, which creates a regular INDEX entry which displays this field's page number,
        // and does not invoke the CrossReferenceSeparator value.
        // The entry for this XE field will display "Apple, 2".
        builder->InsertBreak(BreakType::PageBreak);
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Apple");

        ASSERT_EQ(u" XE  Apple", indexEntry->GetFieldCode());

        // Insert another XE field on page 3 and set a value for the PageNumberReplacement property.
        // This value will show up instead of the number of the page that this field is on,
        // and the INDEX field's CrossReferenceSeparator value will appear in front of it.
        // The entry for this XE field will display "Banana, see: Tropical fruit".
        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Banana");
        indexEntry->set_PageNumberReplacement(u"Tropical fruit");

        ASSERT_EQ(u" XE  Banana \\t \"Tropical fruit\"", indexEntry->GetFieldCode());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.INDEX.XE.CrossReferenceSeparator.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.INDEX.XE.CrossReferenceSeparator.docx");
        index = System::DynamicCast<FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" INDEX  \\k \", see: \"", String(u"Apple, 2\r") + u"Banana, see: Tropical fruit\r", index);
        ASSERT_EQ(u", see: ", index->get_CrossReferenceSeparator());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Apple", String::Empty, indexEntry);
        ASSERT_EQ(u"Apple", indexEntry->get_Text());
        ASSERT_TRUE(indexEntry->get_PageNumberReplacement() == nullptr);

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Banana \\t \"Tropical fruit\"", String::Empty, indexEntry);
        ASSERT_EQ(u"Banana", indexEntry->get_Text());
        ASSERT_EQ(u"Tropical fruit", indexEntry->get_PageNumberReplacement());
    }

    void FieldIndexSubheading(bool runSubentriesOnTheSameLine)
    {
        //ExStart
        //ExFor:FieldIndex.RunSubentriesOnSameLine
        //ExSummary:Shows how to work with subentries in an INDEX field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // The INDEX entry will collect all XE fields with matching values in the "Text" property
        // into one entry as opposed to making an entry for each XE field.
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));
        index->set_PageNumberSeparator(u", see page ");
        index->set_Heading(u"A");

        // XE fields that have a Text property whose value becomes the heading of the INDEX entry.
        // If this value contains two string segments split by a colon (the INDEX entry will treat :) delimiter,
        // the first segment is heading, and the second segment will become the subheading.
        // The INDEX field first groups entries alphabetically, then, if there are multiple XE fields with the same
        // headings, the INDEX field will further subgroup them by the values of these headings.
        // There can be multiple subgrouping layers, depending on how many times
        // the Text properties of XE fields get segmented like this.
        // By default, an INDEX field entry group will create a new line for every subheading within this group.
        // We can set the RunSubentriesOnSameLine flag to true to keep the heading,
        // and every subheading for the group on one line instead, which will make the INDEX field more compact.
        index->set_RunSubentriesOnSameLine(runSubentriesOnTheSameLine);

        if (runSubentriesOnTheSameLine)
        {
            ASSERT_EQ(u" INDEX  \\e \", see page \" \\h A \\r", index->GetFieldCode());
        }
        else
        {
            ASSERT_EQ(u" INDEX  \\e \", see page \" \\h A", index->GetFieldCode());
        }

        // Insert two XE fields, each on a new page, and with the same heading named "Heading 1",
        // which the INDEX field will use to group them.
        // If RunSubentriesOnSameLine is false, then the INDEX table will create three lines:
        // one line for the grouping heading "Heading 1", and one more line for each subheading.
        // If RunSubentriesOnSameLine is true, then the INDEX table will create a one-line
        // entry that encompasses the heading and every subheading.
        builder->InsertBreak(BreakType::PageBreak);
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Heading 1:Subheading 1");

        ASSERT_EQ(u" XE  \"Heading 1:Subheading 1\"", indexEntry->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Heading 1:Subheading 2");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + String::Format(u"Field.INDEX.XE.Subheading.docx"));
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + String::Format(u"Field.INDEX.XE.Subheading.docx"));
        index = System::DynamicCast<FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));

        if (runSubentriesOnTheSameLine)
        {
            TestUtil::VerifyField(FieldType::FieldIndex, u" INDEX  \\r \\e \", see page \" \\h A",
                                  String(u"H\r") + u"Heading 1: Subheading 1, see page 2; Subheading 2, see page 3\r", index);
            ASSERT_TRUE(index->get_RunSubentriesOnSameLine());
        }
        else
        {
            TestUtil::VerifyField(FieldType::FieldIndex, u" INDEX  \\e \", see page \" \\h A",
                                  String(u"H\r") + u"Heading 1\r" + u"Subheading 1, see page 2\r" + u"Subheading 2, see page 3\r", index);
            ASSERT_FALSE(index->get_RunSubentriesOnSameLine());
        }

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  \"Heading 1:Subheading 1\"", String::Empty, indexEntry);
        ASSERT_EQ(u"Heading 1:Subheading 1", indexEntry->get_Text());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  \"Heading 1:Subheading 2\"", String::Empty, indexEntry);
        ASSERT_EQ(u"Heading 1:Subheading 2", indexEntry->get_Text());
    }

    void FieldIndexYomi(bool sortEntriesUsingYomi)
    {
        //ExStart
        //ExFor:FieldIndex.UseYomi
        //ExFor:FieldXE.Yomi
        //ExSummary:Shows how to sort INDEX field entries phonetically.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // The INDEX entry will collect all XE fields with matching values in the "Text" property
        // into one entry as opposed to making an entry for each XE field.
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        // The INDEX table automatically sorts its entries by the values of their Text properties in alphabetic order.
        // Set the INDEX table to sort entries phonetically using Hiragana instead.
        index->set_UseYomi(sortEntriesUsingYomi);

        if (sortEntriesUsingYomi)
        {
            ASSERT_EQ(u" INDEX  \\y", index->GetFieldCode());
        }
        else
        {
            ASSERT_EQ(u" INDEX ", index->GetFieldCode());
        }

        // Insert 4 XE fields, which would show up as entries in the INDEX field's table of contents.
        // The "Text" property may contain a word's spelling in Kanji, whose pronunciation may be ambiguous,
        // while the "Yomi" version of the word will spell exactly how it is pronounced using Hiragana.
        // If we set our INDEX field to use Yomi, it will sort these entries
        // by the value of their Yomi properties, instead of their Text values.
        builder->InsertBreak(BreakType::PageBreak);
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"愛子");
        indexEntry->set_Yomi(u"あ");

        ASSERT_EQ(u" XE  愛子 \\y あ", indexEntry->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"明美");
        indexEntry->set_Yomi(u"あ");

        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"恵美");
        indexEntry->set_Yomi(u"え");

        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"愛美");
        indexEntry->set_Yomi(u"え");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.INDEX.XE.Yomi.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.INDEX.XE.Yomi.docx");
        index = System::DynamicCast<FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));

        if (sortEntriesUsingYomi)
        {
            ASSERT_TRUE(index->get_UseYomi());
            ASSERT_EQ(u" INDEX  \\y", index->GetFieldCode());
            ASSERT_EQ(String(u"愛子, 2\r") + u"明美, 3\r" + u"恵美, 4\r" + u"愛美, 5\r", index->get_Result());
        }
        else
        {
            ASSERT_FALSE(index->get_UseYomi());
            ASSERT_EQ(u" INDEX ", index->GetFieldCode());
            ASSERT_EQ(String(u"恵美, 4\r") + u"愛子, 2\r" + u"愛美, 5\r" + u"明美, 3\r", index->get_Result());
        }

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  愛子 \\y あ", String::Empty, indexEntry);
        ASSERT_EQ(u"愛子", indexEntry->get_Text());
        ASSERT_EQ(u"あ", indexEntry->get_Yomi());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  明美 \\y あ", String::Empty, indexEntry);
        ASSERT_EQ(u"明美", indexEntry->get_Text());
        ASSERT_EQ(u"あ", indexEntry->get_Yomi());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  恵美 \\y え", String::Empty, indexEntry);
        ASSERT_EQ(u"恵美", indexEntry->get_Text());
        ASSERT_EQ(u"え", indexEntry->get_Yomi());

        indexEntry = System::DynamicCast<FieldXE>(doc->get_Range()->get_Fields()->idx_get(4));

        TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  愛美 \\y え", String::Empty, indexEntry);
        ASSERT_EQ(u"愛美", indexEntry->get_Text());
        ASSERT_EQ(u"え", indexEntry->get_Yomi());
    }

    void FieldBarcode_()
    {
        //ExStart
        //ExFor:FieldBarcode
        //ExFor:FieldBarcode.FacingIdentificationMark
        //ExFor:FieldBarcode.IsBookmark
        //ExFor:FieldBarcode.IsUSPostalAddress
        //ExFor:FieldBarcode.PostalAddress
        //ExSummary:Shows how to use the BARCODE field to display U.S. ZIP codes in the form of a barcode.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln();

        // Below are two ways of using BARCODE fields to display custom values as barcodes.
        // 1 -  Store the value that the barcode will display in the PostalAddress property:
        auto field = System::DynamicCast<FieldBarcode>(builder->InsertField(FieldType::FieldBarcode, true));

        // This value needs to be a valid ZIP code.
        field->set_PostalAddress(u"96801");
        field->set_IsUSPostalAddress(true);
        field->set_FacingIdentificationMark(u"C");

        ASSERT_EQ(u" BARCODE  96801 \\u \\f C", field->GetFieldCode());

        builder->InsertBreak(BreakType::LineBreak);

        // 2 -  Reference a bookmark that stores the value that this barcode will display:
        field = System::DynamicCast<FieldBarcode>(builder->InsertField(FieldType::FieldBarcode, true));
        field->set_PostalAddress(u"BarcodeBookmark");
        field->set_IsBookmark(true);

        ASSERT_EQ(u" BARCODE  BarcodeBookmark \\b", field->GetFieldCode());

        // The bookmark that the BARCODE field references in its PostalAddress property
        // need to contain nothing besides the valid ZIP code.
        builder->InsertBreak(BreakType::PageBreak);
        builder->StartBookmark(u"BarcodeBookmark");
        builder->Writeln(u"968877");
        builder->EndBookmark(u"BarcodeBookmark");

        doc->Save(ArtifactsDir + u"Field.BARCODE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.BARCODE.docx");

        ASSERT_EQ(0, doc->GetChildNodes(NodeType::Shape, true)->get_Count());

        field = System::DynamicCast<FieldBarcode>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldBarcode, u" BARCODE  96801 \\u \\f C", String::Empty, field);
        ASSERT_EQ(u"C", field->get_FacingIdentificationMark());
        ASSERT_EQ(u"96801", field->get_PostalAddress());
        ASSERT_TRUE(field->get_IsUSPostalAddress());

        field = System::DynamicCast<FieldBarcode>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldBarcode, u" BARCODE  BarcodeBookmark \\b", String::Empty, field);
        ASSERT_EQ(u"BarcodeBookmark", field->get_PostalAddress());
        ASSERT_TRUE(field->get_IsBookmark());
    }

    void FieldDisplayBarcode_()
    {
        //ExStart
        //ExFor:FieldDisplayBarcode
        //ExFor:FieldDisplayBarcode.AddStartStopChar
        //ExFor:FieldDisplayBarcode.BackgroundColor
        //ExFor:FieldDisplayBarcode.BarcodeType
        //ExFor:FieldDisplayBarcode.BarcodeValue
        //ExFor:FieldDisplayBarcode.CaseCodeStyle
        //ExFor:FieldDisplayBarcode.DisplayText
        //ExFor:FieldDisplayBarcode.ErrorCorrectionLevel
        //ExFor:FieldDisplayBarcode.FixCheckDigit
        //ExFor:FieldDisplayBarcode.ForegroundColor
        //ExFor:FieldDisplayBarcode.PosCodeStyle
        //ExFor:FieldDisplayBarcode.ScalingFactor
        //ExFor:FieldDisplayBarcode.SymbolHeight
        //ExFor:FieldDisplayBarcode.SymbolRotation
        //ExSummary:Shows how to insert a DISPLAYBARCODE field, and set its properties.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto field = System::DynamicCast<FieldDisplayBarcode>(builder->InsertField(FieldType::FieldDisplayBarcode, true));

        // Below are four types of barcodes, decorated in various ways, that the DISPLAYBARCODE field can display.
        // 1 -  QR code with custom colors:
        field->set_BarcodeType(u"QR");
        field->set_BarcodeValue(u"ABC123");
        field->set_BackgroundColor(u"0xF8BD69");
        field->set_ForegroundColor(u"0xB5413B");
        field->set_ErrorCorrectionLevel(u"3");
        field->set_ScalingFactor(u"250");
        field->set_SymbolHeight(u"1000");
        field->set_SymbolRotation(u"0");

        ASSERT_EQ(u" DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0", field->GetFieldCode());
        builder->Writeln();

        // 2 -  EAN13 barcode, with the digits displayed below the bars:
        field = System::DynamicCast<FieldDisplayBarcode>(builder->InsertField(FieldType::FieldDisplayBarcode, true));
        field->set_BarcodeType(u"EAN13");
        field->set_BarcodeValue(u"501234567890");
        field->set_DisplayText(true);
        field->set_PosCodeStyle(u"CASE");
        field->set_FixCheckDigit(true);

        ASSERT_EQ(u" DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", field->GetFieldCode());
        builder->Writeln();

        // 3 -  CODE39 barcode:
        field = System::DynamicCast<FieldDisplayBarcode>(builder->InsertField(FieldType::FieldDisplayBarcode, true));
        field->set_BarcodeType(u"CODE39");
        field->set_BarcodeValue(u"12345ABCDE");
        field->set_AddStartStopChar(true);

        ASSERT_EQ(u" DISPLAYBARCODE  12345ABCDE CODE39 \\d", field->GetFieldCode());
        builder->Writeln();

        // 4 -  ITF4 barcode, with a specified case code:
        field = System::DynamicCast<FieldDisplayBarcode>(builder->InsertField(FieldType::FieldDisplayBarcode, true));
        field->set_BarcodeType(u"ITF14");
        field->set_BarcodeValue(u"09312345678907");
        field->set_CaseCodeStyle(u"STD");

        ASSERT_EQ(u" DISPLAYBARCODE  09312345678907 ITF14 \\c STD", field->GetFieldCode());

        doc->Save(ArtifactsDir + u"Field.DISPLAYBARCODE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.DISPLAYBARCODE.docx");

        ASSERT_EQ(0, doc->GetChildNodes(NodeType::Shape, true)->get_Count());

        field = System::DynamicCast<FieldDisplayBarcode>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldDisplayBarcode, u" DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0",
                              String::Empty, field);
        ASSERT_EQ(u"QR", field->get_BarcodeType());
        ASSERT_EQ(u"ABC123", field->get_BarcodeValue());
        ASSERT_EQ(u"0xF8BD69", field->get_BackgroundColor());
        ASSERT_EQ(u"0xB5413B", field->get_ForegroundColor());
        ASSERT_EQ(u"3", field->get_ErrorCorrectionLevel());
        ASSERT_EQ(u"250", field->get_ScalingFactor());
        ASSERT_EQ(u"1000", field->get_SymbolHeight());
        ASSERT_EQ(u"0", field->get_SymbolRotation());

        field = System::DynamicCast<FieldDisplayBarcode>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldDisplayBarcode, u" DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", String::Empty, field);
        ASSERT_EQ(u"EAN13", field->get_BarcodeType());
        ASSERT_EQ(u"501234567890", field->get_BarcodeValue());
        ASSERT_TRUE(field->get_DisplayText());
        ASSERT_EQ(u"CASE", field->get_PosCodeStyle());
        ASSERT_TRUE(field->get_FixCheckDigit());

        field = System::DynamicCast<FieldDisplayBarcode>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldDisplayBarcode, u" DISPLAYBARCODE  12345ABCDE CODE39 \\d", String::Empty, field);
        ASSERT_EQ(u"CODE39", field->get_BarcodeType());
        ASSERT_EQ(u"12345ABCDE", field->get_BarcodeValue());
        ASSERT_TRUE(field->get_AddStartStopChar());

        field = System::DynamicCast<FieldDisplayBarcode>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldDisplayBarcode, u" DISPLAYBARCODE  09312345678907 ITF14 \\c STD", String::Empty, field);
        ASSERT_EQ(u"ITF14", field->get_BarcodeType());
        ASSERT_EQ(u"09312345678907", field->get_BarcodeValue());
        ASSERT_EQ(u"STD", field->get_CaseCodeStyle());
    }

    //ExStart
    //ExFor:FieldLink
    //ExFor:FieldLink.AutoUpdate
    //ExFor:FieldLink.FormatUpdateType
    //ExFor:FieldLink.InsertAsBitmap
    //ExFor:FieldLink.InsertAsHtml
    //ExFor:FieldLink.InsertAsPicture
    //ExFor:FieldLink.InsertAsRtf
    //ExFor:FieldLink.InsertAsText
    //ExFor:FieldLink.InsertAsUnicode
    //ExFor:FieldLink.IsLinked
    //ExFor:FieldLink.ProgId
    //ExFor:FieldLink.SourceFullName
    //ExFor:FieldLink.SourceItem
    //ExFor:FieldDde
    //ExFor:FieldDde.AutoUpdate
    //ExFor:FieldDde.InsertAsBitmap
    //ExFor:FieldDde.InsertAsHtml
    //ExFor:FieldDde.InsertAsPicture
    //ExFor:FieldDde.InsertAsRtf
    //ExFor:FieldDde.InsertAsText
    //ExFor:FieldDde.InsertAsUnicode
    //ExFor:FieldDde.IsLinked
    //ExFor:FieldDde.ProgId
    //ExFor:FieldDde.SourceFullName
    //ExFor:FieldDde.SourceItem
    //ExFor:FieldDdeAuto
    //ExFor:FieldDdeAuto.InsertAsBitmap
    //ExFor:FieldDdeAuto.InsertAsHtml
    //ExFor:FieldDdeAuto.InsertAsPicture
    //ExFor:FieldDdeAuto.InsertAsRtf
    //ExFor:FieldDdeAuto.InsertAsText
    //ExFor:FieldDdeAuto.InsertAsUnicode
    //ExFor:FieldDdeAuto.IsLinked
    //ExFor:FieldDdeAuto.ProgId
    //ExFor:FieldDdeAuto.SourceFullName
    //ExFor:FieldDdeAuto.SourceItem
    //ExSummary:Shows how to use various field types to link to other documents in the local file system, and display their contents.
    enum class InsertLinkedObjectAs
    {
        Text,
        Unicode,
        Html,
        Rtf,
        Picture,
        Bitmap
    };
    void FieldLinkedObjectsAsText(ExField::InsertLinkedObjectAs insertLinkedObjectAs)
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are three types of fields we can use to display contents from a linked document in the form of text.
        // 1 -  A LINK field:
        builder->Writeln(u"FieldLink:\n");
        InsertFieldLink(builder, insertLinkedObjectAs, u"Word.Document.8", MyDir + u"Document.docx", nullptr, true);

        // 2 -  A DDE field:
        builder->Writeln(u"FieldDde:\n");
        InsertFieldDde(builder, insertLinkedObjectAs, u"Excel.Sheet", MyDir + u"Spreadsheet.xlsx", u"Sheet1!R1C1", true, true);

        // 3 -  A DDEAUTO field:
        builder->Writeln(u"FieldDdeAuto:\n");
        InsertFieldDdeAuto(builder, insertLinkedObjectAs, u"Excel.Sheet", MyDir + u"Spreadsheet.xlsx", u"Sheet1!R1C1", true);

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.LINK.DDE.DDEAUTO.docx");
    }

    void FieldLinkedObjectsAsImage(ExField::InsertLinkedObjectAs insertLinkedObjectAs)
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are three types of fields we can use to display contents from a linked document in the form of an image.
        // 1 -  A LINK field:
        builder->Writeln(u"FieldLink:\n");
        InsertFieldLink(builder, insertLinkedObjectAs, u"Excel.Sheet", MyDir + u"MySpreadsheet.xlsx", u"Sheet1!R2C2", true);

        // 2 -  A DDE field:
        builder->Writeln(u"FieldDde:\n");
        InsertFieldDde(builder, insertLinkedObjectAs, u"Excel.Sheet", MyDir + u"Spreadsheet.xlsx", u"Sheet1!R1C1", true, true);

        // 3 -  A DDEAUTO field:
        builder->Writeln(u"FieldDdeAuto:\n");
        InsertFieldDdeAuto(builder, insertLinkedObjectAs, u"Excel.Sheet", MyDir + u"Spreadsheet.xlsx", u"Sheet1!R1C1", true);

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.LINK.DDE.DDEAUTO.AsImage.docx");
    }

    /// <summary>
    /// Use a document builder to insert a LINK field and set its properties according to parameters.
    /// </summary>
    static void InsertFieldLink(SharedPtr<DocumentBuilder> builder, ExField::InsertLinkedObjectAs insertLinkedObjectAs, String progId, String sourceFullName,
                                String sourceItem, bool shouldAutoUpdate)
    {
        auto field = System::DynamicCast<FieldLink>(builder->InsertField(FieldType::FieldLink, true));

        switch (insertLinkedObjectAs)
        {
        case ApiExamples::ExField::InsertLinkedObjectAs::Text:
            field->set_InsertAsText(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Unicode:
            field->set_InsertAsUnicode(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Html:
            field->set_InsertAsHtml(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Rtf:
            field->set_InsertAsRtf(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Picture:
            field->set_InsertAsPicture(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Bitmap:
            field->set_InsertAsBitmap(true);
            break;
        }

        field->set_AutoUpdate(shouldAutoUpdate);
        field->set_ProgId(progId);
        field->set_SourceFullName(sourceFullName);
        field->set_SourceItem(sourceItem);

        builder->Writeln(u"\n");
    }

    /// <summary>
    /// Use a document builder to insert a DDE field, and set its properties according to parameters.
    /// </summary>
    static void InsertFieldDde(SharedPtr<DocumentBuilder> builder, ExField::InsertLinkedObjectAs insertLinkedObjectAs, String progId, String sourceFullName,
                               String sourceItem, bool isLinked, bool shouldAutoUpdate)
    {
        auto field = System::DynamicCast<FieldDde>(builder->InsertField(FieldType::FieldDDE, true));

        switch (insertLinkedObjectAs)
        {
        case ApiExamples::ExField::InsertLinkedObjectAs::Text:
            field->set_InsertAsText(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Unicode:
            field->set_InsertAsUnicode(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Html:
            field->set_InsertAsHtml(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Rtf:
            field->set_InsertAsRtf(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Picture:
            field->set_InsertAsPicture(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Bitmap:
            field->set_InsertAsBitmap(true);
            break;
        }

        field->set_AutoUpdate(shouldAutoUpdate);
        field->set_ProgId(progId);
        field->set_SourceFullName(sourceFullName);
        field->set_SourceItem(sourceItem);
        field->set_IsLinked(isLinked);

        builder->Writeln(u"\n");
    }

    /// <summary>
    /// Use a document builder to insert a DDEAUTO, field and set its properties according to parameters.
    /// </summary>
    static void InsertFieldDdeAuto(SharedPtr<DocumentBuilder> builder, ExField::InsertLinkedObjectAs insertLinkedObjectAs, String progId, String sourceFullName,
                                   String sourceItem, bool isLinked)
    {
        auto field = System::DynamicCast<FieldDdeAuto>(builder->InsertField(FieldType::FieldDDEAuto, true));

        switch (insertLinkedObjectAs)
        {
        case ApiExamples::ExField::InsertLinkedObjectAs::Text:
            field->set_InsertAsText(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Unicode:
            field->set_InsertAsUnicode(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Html:
            field->set_InsertAsHtml(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Rtf:
            field->set_InsertAsRtf(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Picture:
            field->set_InsertAsPicture(true);
            break;

        case ApiExamples::ExField::InsertLinkedObjectAs::Bitmap:
            field->set_InsertAsBitmap(true);
            break;
        }

        field->set_ProgId(progId);
        field->set_SourceFullName(sourceFullName);
        field->set_SourceItem(sourceItem);
        field->set_IsLinked(isLinked);
    }

    //ExEnd

    void FieldUserAddress_()
    {
        //ExStart
        //ExFor:FieldUserAddress
        //ExFor:FieldUserAddress.UserAddress
        //ExSummary:Shows how to use the USERADDRESS field.
        auto doc = MakeObject<Document>();

        // Create a UserInformation object and set it as the source of user information for any fields that we create.
        auto userInformation = MakeObject<UserInformation>();
        userInformation->set_Address(u"123 Main Street");
        doc->get_FieldOptions()->set_CurrentUser(userInformation);

        // Create a USERADDRESS field to display the current user's address,
        // taken from the UserInformation object we created above.
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto fieldUserAddress = System::DynamicCast<FieldUserAddress>(builder->InsertField(FieldType::FieldUserAddress, true));
        ASSERT_EQ(userInformation->get_Address(), fieldUserAddress->get_Result());
        //ExSkip

        ASSERT_EQ(u" USERADDRESS ", fieldUserAddress->GetFieldCode());
        ASSERT_EQ(u"123 Main Street", fieldUserAddress->get_Result());

        // We can set this property to get our field to override the value currently stored in the UserInformation object.
        fieldUserAddress->set_UserAddress(u"456 North Road");
        fieldUserAddress->Update();

        ASSERT_EQ(u" USERADDRESS  \"456 North Road\"", fieldUserAddress->GetFieldCode());
        ASSERT_EQ(u"456 North Road", fieldUserAddress->get_Result());

        // This does not affect the value in the UserInformation object.
        ASSERT_EQ(u"123 Main Street", doc->get_FieldOptions()->get_CurrentUser()->get_Address());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.USERADDRESS.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.USERADDRESS.docx");

        fieldUserAddress = System::DynamicCast<FieldUserAddress>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldUserAddress, u" USERADDRESS  \"456 North Road\"", u"456 North Road", fieldUserAddress);
        ASSERT_EQ(u"456 North Road", fieldUserAddress->get_UserAddress());
    }

    void FieldUserInitials_()
    {
        //ExStart
        //ExFor:FieldUserInitials
        //ExFor:FieldUserInitials.UserInitials
        //ExSummary:Shows how to use the USERINITIALS field.
        auto doc = MakeObject<Document>();

        // Create a UserInformation object and set it as the source of user information for any fields that we create.
        auto userInformation = MakeObject<UserInformation>();
        userInformation->set_Initials(u"J. D.");
        doc->get_FieldOptions()->set_CurrentUser(userInformation);

        // Create a USERINITIALS field to display the current user's initials,
        // taken from the UserInformation object we created above.
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto fieldUserInitials = System::DynamicCast<FieldUserInitials>(builder->InsertField(FieldType::FieldUserInitials, true));
        ASSERT_EQ(userInformation->get_Initials(), fieldUserInitials->get_Result());

        ASSERT_EQ(u" USERINITIALS ", fieldUserInitials->GetFieldCode());
        ASSERT_EQ(u"J. D.", fieldUserInitials->get_Result());

        // We can set this property to get our field to override the value currently stored in the UserInformation object.
        fieldUserInitials->set_UserInitials(u"J. C.");
        fieldUserInitials->Update();

        ASSERT_EQ(u" USERINITIALS  \"J. C.\"", fieldUserInitials->GetFieldCode());
        ASSERT_EQ(u"J. C.", fieldUserInitials->get_Result());

        // This does not affect the value in the UserInformation object.
        ASSERT_EQ(u"J. D.", doc->get_FieldOptions()->get_CurrentUser()->get_Initials());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.USERINITIALS.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.USERINITIALS.docx");

        fieldUserInitials = System::DynamicCast<FieldUserInitials>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldUserInitials, u" USERINITIALS  \"J. C.\"", u"J. C.", fieldUserInitials);
        ASSERT_EQ(u"J. C.", fieldUserInitials->get_UserInitials());
    }

    void FieldUserName_()
    {
        //ExStart
        //ExFor:FieldUserName
        //ExFor:FieldUserName.UserName
        //ExSummary:Shows how to use the USERNAME field.
        auto doc = MakeObject<Document>();

        // Create a UserInformation object and set it as the source of user information for any fields that we create.
        auto userInformation = MakeObject<UserInformation>();
        userInformation->set_Name(u"John Doe");
        doc->get_FieldOptions()->set_CurrentUser(userInformation);

        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a USERNAME field to display the current user's name,
        // taken from the UserInformation object we created above.
        auto fieldUserName = System::DynamicCast<FieldUserName>(builder->InsertField(FieldType::FieldUserName, true));
        ASSERT_EQ(userInformation->get_Name(), fieldUserName->get_Result());

        ASSERT_EQ(u" USERNAME ", fieldUserName->GetFieldCode());
        ASSERT_EQ(u"John Doe", fieldUserName->get_Result());

        // We can set this property to get our field to override the value currently stored in the UserInformation object.
        fieldUserName->set_UserName(u"Jane Doe");
        fieldUserName->Update();

        ASSERT_EQ(u" USERNAME  \"Jane Doe\"", fieldUserName->GetFieldCode());
        ASSERT_EQ(u"Jane Doe", fieldUserName->get_Result());

        // This does not affect the value in the UserInformation object.
        ASSERT_EQ(u"John Doe", doc->get_FieldOptions()->get_CurrentUser()->get_Name());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.USERNAME.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.USERNAME.docx");

        fieldUserName = System::DynamicCast<FieldUserName>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldUserName, u" USERNAME  \"Jane Doe\"", u"Jane Doe", fieldUserName);
        ASSERT_EQ(u"Jane Doe", fieldUserName->get_UserName());
    }

    void FieldStyleRefParagraphNumbers()
    {
        //ExStart
        //ExFor:FieldStyleRef
        //ExFor:FieldStyleRef.InsertParagraphNumber
        //ExFor:FieldStyleRef.InsertParagraphNumberInFullContext
        //ExFor:FieldStyleRef.InsertParagraphNumberInRelativeContext
        //ExFor:FieldStyleRef.InsertRelativePosition
        //ExFor:FieldStyleRef.SearchFromBottom
        //ExFor:FieldStyleRef.StyleName
        //ExFor:FieldStyleRef.SuppressNonDelimiters
        //ExSummary:Shows how to use STYLEREF fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a list based using a Microsoft Word list template.
        SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(Lists::ListTemplate::NumberDefault);

        // This generated list will display "1.a )".
        // Space before the bracket is a non-delimiter character, which we can suppress.
        list->get_ListLevels()->idx_get(0)->set_NumberFormat(u"\x0000"
                                                             u".");
        list->get_ListLevels()->idx_get(1)->set_NumberFormat(u"\x0001"
                                                             u" )");

        // Add text and apply paragraph styles that STYLEREF fields will reference.
        builder->get_ListFormat()->set_List(list);
        builder->get_ListFormat()->ListIndent();
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"List Paragraph"));
        builder->Writeln(u"Item 1");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Quote"));
        builder->Writeln(u"Item 2");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"List Paragraph"));
        builder->Writeln(u"Item 3");
        builder->get_ListFormat()->RemoveNumbers();
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Normal"));

        // Place a STYLEREF field in the header and display the first "List Paragraph"-styled text in the document.
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        auto field = System::DynamicCast<FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
        field->set_StyleName(u"List Paragraph");

        // Place a STYLEREF field in the footer, and have it display the last text.
        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        field = System::DynamicCast<FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
        field->set_StyleName(u"List Paragraph");
        field->set_SearchFromBottom(true);

        builder->MoveToDocumentEnd();

        // We can also use STYLEREF fields to reference the list numbers of lists.
        builder->Write(u"\nParagraph number: ");
        field = System::DynamicCast<FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
        field->set_StyleName(u"Quote");
        field->set_InsertParagraphNumber(true);

        builder->Write(u"\nParagraph number, relative context: ");
        field = System::DynamicCast<FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
        field->set_StyleName(u"Quote");
        field->set_InsertParagraphNumberInRelativeContext(true);

        builder->Write(u"\nParagraph number, full context: ");
        field = System::DynamicCast<FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
        field->set_StyleName(u"Quote");
        field->set_InsertParagraphNumberInFullContext(true);

        builder->Write(u"\nParagraph number, full context, non-delimiter chars suppressed: ");
        field = System::DynamicCast<FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
        field->set_StyleName(u"Quote");
        field->set_InsertParagraphNumberInFullContext(true);
        field->set_SuppressNonDelimiters(true);

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.STYLEREF.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.STYLEREF.docx");

        field = System::DynamicCast<FieldStyleRef>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldStyleRef, u" STYLEREF  \"List Paragraph\"", u"Item 1", field);
        ASSERT_EQ(u"List Paragraph", field->get_StyleName());

        field = System::DynamicCast<FieldStyleRef>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldStyleRef, u" STYLEREF  \"List Paragraph\" \\l", u"Item 3", field);
        ASSERT_EQ(u"List Paragraph", field->get_StyleName());
        ASSERT_TRUE(field->get_SearchFromBottom());

        field = System::DynamicCast<FieldStyleRef>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldStyleRef, u" STYLEREF  Quote \\n", u"b )", field);
        ASSERT_EQ(u"Quote", field->get_StyleName());
        ASSERT_TRUE(field->get_InsertParagraphNumber());

        field = System::DynamicCast<FieldStyleRef>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldStyleRef, u" STYLEREF  Quote \\r", u"b )", field);
        ASSERT_EQ(u"Quote", field->get_StyleName());
        ASSERT_TRUE(field->get_InsertParagraphNumberInRelativeContext());

        field = System::DynamicCast<FieldStyleRef>(doc->get_Range()->get_Fields()->idx_get(4));

        TestUtil::VerifyField(FieldType::FieldStyleRef, u" STYLEREF  Quote \\w", u"1.b )", field);
        ASSERT_EQ(u"Quote", field->get_StyleName());
        ASSERT_TRUE(field->get_InsertParagraphNumberInFullContext());

        field = System::DynamicCast<FieldStyleRef>(doc->get_Range()->get_Fields()->idx_get(5));

        TestUtil::VerifyField(FieldType::FieldStyleRef, u" STYLEREF  Quote \\w \\t", u"1.b)", field);
        ASSERT_EQ(u"Quote", field->get_StyleName());
        ASSERT_TRUE(field->get_InsertParagraphNumberInFullContext());
        ASSERT_TRUE(field->get_SuppressNonDelimiters());
    }

    void FieldDate_()
    {
        //ExStart
        //ExFor:FieldDate
        //ExFor:FieldDate.UseLunarCalendar
        //ExFor:FieldDate.UseSakaEraCalendar
        //ExFor:FieldDate.UseUmAlQuraCalendar
        //ExFor:FieldDate.UseLastFormat
        //ExSummary:Shows how to use DATE fields to display dates according to different kinds of calendars.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // If we want the text in the document always to display the correct date, we can use a DATE field.
        // Below are three types of cultural calendars that a DATE field can use to display a date.
        // 1 -  Islamic Lunar Calendar:
        auto field = System::DynamicCast<FieldDate>(builder->InsertField(FieldType::FieldDate, true));
        field->set_UseLunarCalendar(true);
        ASSERT_EQ(u" DATE  \\h", field->GetFieldCode());
        builder->Writeln();

        // 2 -  Umm al-Qura calendar:
        field = System::DynamicCast<FieldDate>(builder->InsertField(FieldType::FieldDate, true));
        field->set_UseUmAlQuraCalendar(true);
        ASSERT_EQ(u" DATE  \\u", field->GetFieldCode());
        builder->Writeln();

        // 3 -  Indian National Calendar:
        field = System::DynamicCast<FieldDate>(builder->InsertField(FieldType::FieldDate, true));
        field->set_UseSakaEraCalendar(true);
        ASSERT_EQ(u" DATE  \\s", field->GetFieldCode());
        builder->Writeln();

        // Insert a DATE field and set its calendar type to the one last used by the host application.
        // In Microsoft Word, the type will be the most recently used in the Insert -> Text -> Date and Time dialog box.
        field = System::DynamicCast<FieldDate>(builder->InsertField(FieldType::FieldDate, true));
        field->set_UseLastFormat(true);
        ASSERT_EQ(u" DATE  \\l", field->GetFieldCode());
        builder->Writeln();

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.DATE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.DATE.docx");

        field = System::DynamicCast<FieldDate>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(FieldType::FieldDate, field->get_Type());
        ASSERT_TRUE(field->get_UseLunarCalendar());
        ASSERT_EQ(u" DATE  \\h", field->GetFieldCode());
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(doc->get_Range()->get_Fields()->idx_get(0)->get_Result(), u"\\d{1,2}[/]\\d{1,2}[/]\\d{4}")
                        ->get_Success());

        field = System::DynamicCast<FieldDate>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldDate, u" DATE  \\u", System::DateTime::get_Now().ToShortDateString(), field);
        ASSERT_TRUE(field->get_UseUmAlQuraCalendar());

        field = System::DynamicCast<FieldDate>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldDate, u" DATE  \\s", System::DateTime::get_Now().ToShortDateString(), field);
        ASSERT_TRUE(field->get_UseSakaEraCalendar());

        field = System::DynamicCast<FieldDate>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldDate, u" DATE  \\l", System::DateTime::get_Now().ToShortDateString(), field);
        ASSERT_TRUE(field->get_UseLastFormat());
    }

    void FieldCreateDate_()
    {
        //ExStart
        //ExFor:FieldCreateDate
        //ExFor:FieldCreateDate.UseLunarCalendar
        //ExFor:FieldCreateDate.UseSakaEraCalendar
        //ExFor:FieldCreateDate.UseUmAlQuraCalendar
        //ExSummary:Shows how to use the CREATEDATE field to display the creation date/time of the document.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveToDocumentEnd();
        builder->Writeln(u" Date this document was created:");

        // We can use the CREATEDATE field to display the date and time of the creation of the document.
        // Below are three different calendar types according to which the CREATEDATE field can display the date/time.
        // 1 -  Islamic Lunar Calendar:
        builder->Write(u"According to the Lunar Calendar - ");
        auto field = System::DynamicCast<FieldCreateDate>(builder->InsertField(FieldType::FieldCreateDate, true));
        field->set_UseLunarCalendar(true);

        ASSERT_EQ(u" CREATEDATE  \\h", field->GetFieldCode());

        // 2 -  Umm al-Qura calendar:
        builder->Write(u"\nAccording to the Umm al-Qura Calendar - ");
        field = System::DynamicCast<FieldCreateDate>(builder->InsertField(FieldType::FieldCreateDate, true));
        field->set_UseUmAlQuraCalendar(true);

        ASSERT_EQ(u" CREATEDATE  \\u", field->GetFieldCode());

        // 3 -  Indian National Calendar:
        builder->Write(u"\nAccording to the Indian National Calendar - ");
        field = System::DynamicCast<FieldCreateDate>(builder->InsertField(FieldType::FieldCreateDate, true));
        field->set_UseSakaEraCalendar(true);

        ASSERT_EQ(u" CREATEDATE  \\s", field->GetFieldCode());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.CREATEDATE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.CREATEDATE.docx");

        ASSERT_EQ(System::DateTime(2017, 12, 5, 9, 56, 0), doc->get_BuiltInDocumentProperties()->get_CreatedTime());

        System::DateTime expectedDate = doc->get_BuiltInDocumentProperties()->get_CreatedTime().AddHours(
            System::TimeZoneInfo::get_Local()->GetUtcOffset(System::DateTime::get_UtcNow()).get_Hours());
        field = System::DynamicCast<FieldCreateDate>(doc->get_Range()->get_Fields()->idx_get(0));
        SharedPtr<System::Globalization::Calendar> umAlQuraCalendar = MakeObject<System::Globalization::UmAlQuraCalendar>();

        TestUtil::VerifyField(FieldType::FieldCreateDate, u" CREATEDATE  \\h",
                              String::Format(u"{0}/{1}/{2} ", umAlQuraCalendar->GetMonth(expectedDate), umAlQuraCalendar->GetDayOfMonth(expectedDate),
                                             umAlQuraCalendar->GetYear(expectedDate)) +
                                  expectedDate.AddHours(1).ToString(u"hh:mm:ss tt"),
                              field);
        ASSERT_EQ(FieldType::FieldCreateDate, field->get_Type());
        ASSERT_TRUE(field->get_UseLunarCalendar());

        field = System::DynamicCast<FieldCreateDate>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldCreateDate, u" CREATEDATE  \\u",
                              String::Format(u"{0}/{1}/{2} ", umAlQuraCalendar->GetMonth(expectedDate), umAlQuraCalendar->GetDayOfMonth(expectedDate),
                                             umAlQuraCalendar->GetYear(expectedDate)) +
                                  expectedDate.AddHours(1).ToString(u"hh:mm:ss tt"),
                              field);
        ASSERT_EQ(FieldType::FieldCreateDate, field->get_Type());
        ASSERT_TRUE(field->get_UseUmAlQuraCalendar());
    }

    void FieldSaveDate_()
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.LastSavedTime
        //ExFor:FieldSaveDate
        //ExFor:FieldSaveDate.UseLunarCalendar
        //ExFor:FieldSaveDate.UseSakaEraCalendar
        //ExFor:FieldSaveDate.UseUmAlQuraCalendar
        //ExSummary:Shows how to use the SAVEDATE field to display the date/time of the document's most recent save operation performed using Microsoft Word.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveToDocumentEnd();
        builder->Writeln(u" Date this document was last saved:");

        // We can use the SAVEDATE field to display the last save operation's date and time on the document.
        // The save operation that these fields refer to is the manual save in an application like Microsoft Word,
        // not the document's Save method.
        // Below are three different calendar types according to which the SAVEDATE field can display the date/time.
        // 1 -  Islamic Lunar Calendar:
        builder->Write(u"According to the Lunar Calendar - ");
        auto field = System::DynamicCast<FieldSaveDate>(builder->InsertField(FieldType::FieldSaveDate, true));
        field->set_UseLunarCalendar(true);

        ASSERT_EQ(u" SAVEDATE  \\h", field->GetFieldCode());

        // 2 -  Umm al-Qura calendar:
        builder->Write(u"\nAccording to the Umm al-Qura calendar - ");
        field = System::DynamicCast<FieldSaveDate>(builder->InsertField(FieldType::FieldSaveDate, true));
        field->set_UseUmAlQuraCalendar(true);

        ASSERT_EQ(u" SAVEDATE  \\u", field->GetFieldCode());

        // 3 -  Indian National calendar:
        builder->Write(u"\nAccording to the Indian National calendar - ");
        field = System::DynamicCast<FieldSaveDate>(builder->InsertField(FieldType::FieldSaveDate, true));
        field->set_UseSakaEraCalendar(true);

        ASSERT_EQ(u" SAVEDATE  \\s", field->GetFieldCode());

        // The SAVEDATE fields draw their date/time values from the LastSavedTime built-in property.
        // The document's Save method will not update this value, but we can still update it manually.
        doc->get_BuiltInDocumentProperties()->set_LastSavedTime(System::DateTime::get_Now());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.SAVEDATE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.SAVEDATE.docx");

        std::cout << doc->get_BuiltInDocumentProperties()->get_LastSavedTime() << std::endl;

        field = System::DynamicCast<FieldSaveDate>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(FieldType::FieldSaveDate, field->get_Type());
        ASSERT_TRUE(field->get_UseLunarCalendar());
        ASSERT_EQ(u" SAVEDATE  \\h", field->GetFieldCode());

        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(field->get_Result(), u"\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M")
                        ->get_Success());

        field = System::DynamicCast<FieldSaveDate>(doc->get_Range()->get_Fields()->idx_get(1));

        ASSERT_EQ(FieldType::FieldSaveDate, field->get_Type());
        ASSERT_TRUE(field->get_UseUmAlQuraCalendar());
        ASSERT_EQ(u" SAVEDATE  \\u", field->GetFieldCode());
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(field->get_Result(), u"\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M")
                        ->get_Success());
    }

    void FieldBuilder_()
    {
        //ExStart
        //ExFor:FieldBuilder
        //ExFor:FieldBuilder.AddArgument(Int32)
        //ExFor:FieldBuilder.AddArgument(FieldArgumentBuilder)
        //ExFor:FieldBuilder.AddArgument(String)
        //ExFor:FieldBuilder.AddArgument(Double)
        //ExFor:FieldBuilder.AddArgument(FieldBuilder)
        //ExFor:FieldBuilder.AddSwitch(String)
        //ExFor:FieldBuilder.AddSwitch(String, Double)
        //ExFor:FieldBuilder.AddSwitch(String, Int32)
        //ExFor:FieldBuilder.AddSwitch(String, String)
        //ExFor:FieldBuilder.BuildAndInsert(Paragraph)
        //ExFor:FieldArgumentBuilder
        //ExFor:FieldArgumentBuilder.AddField(FieldBuilder)
        //ExFor:FieldArgumentBuilder.AddText(String)
        //ExFor:FieldArgumentBuilder.AddNode(Inline)
        //ExSummary:Shows how to construct fields using a field builder, and then insert them into the document.
        auto doc = MakeObject<Document>();

        // Below are three examples of field construction done using a field builder.
        // 1 -  Single field:
        // Use a field builder to add a SYMBOL field which displays the ƒ (Florin) symbol.
        auto builder = MakeObject<FieldBuilder>(FieldType::FieldSymbol);
        builder->AddArgument(402);
        builder->AddSwitch(u"\\f", u"Arial");
        builder->AddSwitch(u"\\s", 25);
        builder->AddSwitch(u"\\u");
        SharedPtr<Field> field = builder->BuildAndInsert(doc->get_FirstSection()->get_Body()->get_FirstParagraph());

        ASSERT_EQ(u" SYMBOL 402 \\f Arial \\s 25 \\u ", field->GetFieldCode());

        // 2 -  Nested field:
        // Use a field builder to create a formula field used as an inner field by another field builder.
        auto innerFormulaBuilder = MakeObject<FieldBuilder>(FieldType::FieldFormula);
        innerFormulaBuilder->AddArgument(100);
        innerFormulaBuilder->AddArgument(u"+");
        innerFormulaBuilder->AddArgument(74);

        // Create another builder for another SYMBOL field, and insert the formula field
        // that we have created above into the SYMBOL field as its argument.
        builder = MakeObject<FieldBuilder>(FieldType::FieldSymbol);
        builder->AddArgument(innerFormulaBuilder);
        field = builder->BuildAndInsert(doc->get_FirstSection()->get_Body()->AppendParagraph(String::Empty));

        // The outer SYMBOL field will use the formula field result, 174, as its argument,
        // which will make the field display the ® (Registered Sign) symbol since its character number is 174.
        ASSERT_EQ(u" SYMBOL \u0013 = 100 + 74 \u0014\u0015 ", field->GetFieldCode());

        // 3 -  Multiple nested fields and arguments:
        // Now, we will use a builder to create an IF field, which displays one of two custom string values,
        // depending on the true/false value of its expression. To get a true/false value
        // that determines which string the IF field displays, the IF field will test two numeric expressions for equality.
        // We will provide the two expressions in the form of formula fields, which we will nest inside the IF field.
        auto leftExpression = MakeObject<FieldBuilder>(FieldType::FieldFormula);
        leftExpression->AddArgument(2);
        leftExpression->AddArgument(u"+");
        leftExpression->AddArgument(3);

        auto rightExpression = MakeObject<FieldBuilder>(FieldType::FieldFormula);
        rightExpression->AddArgument(2.5);
        rightExpression->AddArgument(u"*");
        rightExpression->AddArgument(5.2);

        // Next, we will build two field arguments, which will serve as the true/false output strings for the IF field.
        // These arguments will reuse the output values of our numeric expressions.
        auto trueOutput = MakeObject<FieldArgumentBuilder>();
        trueOutput->AddText(u"True, both expressions amount to ");
        trueOutput->AddField(leftExpression);

        auto falseOutput = MakeObject<FieldArgumentBuilder>();
        falseOutput->AddNode(MakeObject<Run>(doc, u"False, "));
        falseOutput->AddField(leftExpression);
        falseOutput->AddNode(MakeObject<Run>(doc, u" does not equal "));
        falseOutput->AddField(rightExpression);

        // Finally, we will create one more field builder for the IF field and combine all of the expressions.
        builder = MakeObject<FieldBuilder>(FieldType::FieldIf);
        builder->AddArgument(leftExpression);
        builder->AddArgument(u"=");
        builder->AddArgument(rightExpression);
        builder->AddArgument(trueOutput);
        builder->AddArgument(falseOutput);
        field = builder->BuildAndInsert(doc->get_FirstSection()->get_Body()->AppendParagraph(String::Empty));

        ASSERT_EQ(String(u" IF \u0013 = 2 + 3 \u0014\u0015 = \u0013 = 2.5 * 5.2 \u0014\u0015 ") +
                      u"\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " +
                      u"\"False, \u0013 = 2 + 3 \u0014\u0015 does not equal \u0013 = 2.5 * 5.2 \u0014\u0015\" ",
                  field->GetFieldCode());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.SYMBOL.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.SYMBOL.docx");

        auto fieldSymbol = System::DynamicCast<FieldSymbol>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldSymbol, u" SYMBOL 402 \\f Arial \\s 25 \\u ", String::Empty, fieldSymbol);
        ASSERT_EQ(u"ƒ", fieldSymbol->get_DisplayResult());

        fieldSymbol = System::DynamicCast<FieldSymbol>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldSymbol, u" SYMBOL \u0013 = 100 + 74 \u0014174\u0015 ", String::Empty, fieldSymbol);
        ASSERT_EQ(u"®", fieldSymbol->get_DisplayResult());

        TestUtil::VerifyField(FieldType::FieldFormula, u" = 100 + 74 ", u"174", doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldIf,
                              String(u" IF \u0013 = 2 + 3 \u00145\u0015 = \u0013 = 2.5 * 5.2 \u001413\u0015 ") +
                                  u"\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " +
                                  u"\"False, \u0013 = 2 + 3 \u00145\u0015 does not equal \u0013 = 2.5 * 5.2 \u001413\u0015\" ",
                              u"False, 5 does not equal 13", doc->get_Range()->get_Fields()->idx_get(3));

        ASSERT_NE(doc->get_Range()->get_Fields()->idx_get(2)->get_Start()->get_ParentNode(),
                  doc->get_Range()->get_Fields()->idx_get(3)->get_Start()->get_ParentNode());

        TestUtil::VerifyField(FieldType::FieldFormula, u" = 2 + 3 ", u"5", doc->get_Range()->get_Fields()->idx_get(4));
        TestUtil::FieldsAreNested(doc->get_Range()->get_Fields()->idx_get(4), doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldFormula, u" = 2.5 * 5.2 ", u"13", doc->get_Range()->get_Fields()->idx_get(5));
        TestUtil::FieldsAreNested(doc->get_Range()->get_Fields()->idx_get(5), doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldFormula, u" = 2 + 3 ", String::Empty, doc->get_Range()->get_Fields()->idx_get(6));
        TestUtil::FieldsAreNested(doc->get_Range()->get_Fields()->idx_get(6), doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldFormula, u" = 2 + 3 ", u"5", doc->get_Range()->get_Fields()->idx_get(7));
        TestUtil::FieldsAreNested(doc->get_Range()->get_Fields()->idx_get(7), doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldFormula, u" = 2.5 * 5.2 ", u"13", doc->get_Range()->get_Fields()->idx_get(8));
        TestUtil::FieldsAreNested(doc->get_Range()->get_Fields()->idx_get(8), doc->get_Range()->get_Fields()->idx_get(3));
    }

    void FieldAuthor_()
    {
        //ExStart
        //ExFor:FieldAuthor
        //ExFor:FieldAuthor.AuthorName
        //ExFor:FieldOptions.DefaultDocumentAuthor
        //ExSummary:Shows how to use an AUTHOR field to display a document creator's name.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // AUTHOR fields source their results from the built-in document property called "Author".
        // If we create and save a document in Microsoft Word,
        // it will have our username in that property.
        // However, if we create a document programmatically using Aspose.Words,
        // the "Author" property, by default, will be an empty string.
        ASSERT_EQ(String::Empty, doc->get_BuiltInDocumentProperties()->get_Author());

        // Set a backup author name for AUTHOR fields to use
        // if the "Author" property contains an empty string.
        doc->get_FieldOptions()->set_DefaultDocumentAuthor(u"Joe Bloggs");

        builder->Write(u"This document was created by ");
        auto field = System::DynamicCast<FieldAuthor>(builder->InsertField(FieldType::FieldAuthor, true));
        field->Update();

        ASSERT_EQ(u" AUTHOR ", field->GetFieldCode());
        ASSERT_EQ(u"Joe Bloggs", field->get_Result());

        // Updating an AUTHOR field that contains a value
        // will apply that value to the "Author" built-in property.
        ASSERT_EQ(u"Joe Bloggs", doc->get_BuiltInDocumentProperties()->get_Author());

        // Changing this property, then updating the AUTHOR field will apply this value to the field.
        doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
        field->Update();

        ASSERT_EQ(u" AUTHOR ", field->GetFieldCode());
        ASSERT_EQ(u"John Doe", field->get_Result());

        // If we update an AUTHOR field after changing its "Name" property,
        // then the field will display the new name and apply the new name to the built-in property.
        field->set_AuthorName(u"Jane Doe");
        field->Update();

        ASSERT_EQ(u" AUTHOR  \"Jane Doe\"", field->GetFieldCode());
        ASSERT_EQ(u"Jane Doe", field->get_Result());

        // AUTHOR fields do not affect the DefaultDocumentAuthor property.
        ASSERT_EQ(u"Jane Doe", doc->get_BuiltInDocumentProperties()->get_Author());
        ASSERT_EQ(u"Joe Bloggs", doc->get_FieldOptions()->get_DefaultDocumentAuthor());

        doc->Save(ArtifactsDir + u"Field.AUTHOR.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.AUTHOR.docx");

        ASSERT_TRUE(doc->get_FieldOptions()->get_DefaultDocumentAuthor() == nullptr);
        ASSERT_EQ(u"Jane Doe", doc->get_BuiltInDocumentProperties()->get_Author());

        field = System::DynamicCast<FieldAuthor>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldAuthor, u" AUTHOR  \"Jane Doe\"", u"Jane Doe", field);
        ASSERT_EQ(u"Jane Doe", field->get_AuthorName());
    }

    void FieldDocVariable_()
    {
        //ExStart
        //ExFor:FieldDocProperty
        //ExFor:FieldDocVariable
        //ExFor:FieldDocVariable.VariableName
        //ExSummary:Shows how to use DOCPROPERTY fields to display document properties and variables.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are two ways of using DOCPROPERTY fields.
        // 1 -  Display a built-in property:
        // Set a custom value for the "Category" built-in property, then insert a DOCPROPERTY field that references it.
        doc->get_BuiltInDocumentProperties()->set_Category(u"My category");

        auto fieldDocProperty = System::DynamicCast<FieldDocProperty>(builder->InsertField(u" DOCPROPERTY Category "));
        fieldDocProperty->Update();

        ASSERT_EQ(u" DOCPROPERTY Category ", fieldDocProperty->GetFieldCode());
        ASSERT_EQ(u"My category", fieldDocProperty->get_Result());

        builder->InsertParagraph();

        // 2 -  Display a custom document variable:
        // Define a custom variable, then reference that variable with a DOCPROPERTY field.
        ASSERT_EQ(0, doc->get_Variables()->get_Count());
        doc->get_Variables()->Add(u"My variable", u"My variable's value");

        auto fieldDocVariable = System::DynamicCast<FieldDocVariable>(builder->InsertField(FieldType::FieldDocVariable, true));
        fieldDocVariable->set_VariableName(u"My Variable");
        fieldDocVariable->Update();

        ASSERT_EQ(u" DOCVARIABLE  \"My Variable\"", fieldDocVariable->GetFieldCode());
        ASSERT_EQ(u"My variable's value", fieldDocVariable->get_Result());

        doc->Save(ArtifactsDir + u"Field.DOCPROPERTY.DOCVARIABLE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.DOCPROPERTY.DOCVARIABLE.docx");

        ASSERT_EQ(u"My category", doc->get_BuiltInDocumentProperties()->get_Category());

        fieldDocProperty = System::DynamicCast<FieldDocProperty>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldDocProperty, u" DOCPROPERTY Category ", u"My category", fieldDocProperty);

        fieldDocVariable = System::DynamicCast<FieldDocVariable>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldDocVariable, u" DOCVARIABLE  \"My Variable\"", u"My variable's value", fieldDocVariable);
        ASSERT_EQ(u"My Variable", fieldDocVariable->get_VariableName());
    }

    void FieldSubject_()
    {
        //ExStart
        //ExFor:FieldSubject
        //ExFor:FieldSubject.Text
        //ExSummary:Shows how to use the SUBJECT field.
        auto doc = MakeObject<Document>();

        // Set a value for the document's "Subject" built-in property.
        doc->get_BuiltInDocumentProperties()->set_Subject(u"My subject");

        // Create a SUBJECT field to display the value of that built-in property.
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto field = System::DynamicCast<FieldSubject>(builder->InsertField(FieldType::FieldSubject, true));
        field->Update();

        ASSERT_EQ(u" SUBJECT ", field->GetFieldCode());
        ASSERT_EQ(u"My subject", field->get_Result());

        // If we give the SUBJECT field's Text property value and update it, the field will
        // overwrite the current value of the "Subject" built-in property with the value of its Text property,
        // and then display the new value.
        field->set_Text(u"My new subject");
        field->Update();

        ASSERT_EQ(u" SUBJECT  \"My new subject\"", field->GetFieldCode());
        ASSERT_EQ(u"My new subject", field->get_Result());

        ASSERT_EQ(u"My new subject", doc->get_BuiltInDocumentProperties()->get_Subject());

        doc->Save(ArtifactsDir + u"Field.SUBJECT.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.SUBJECT.docx");

        ASSERT_EQ(u"My new subject", doc->get_BuiltInDocumentProperties()->get_Subject());

        field = System::DynamicCast<FieldSubject>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldSubject, u" SUBJECT  \"My new subject\"", u"My new subject", field);
        ASSERT_EQ(u"My new subject", field->get_Text());
    }

    void FieldComments_()
    {
        //ExStart
        //ExFor:FieldComments
        //ExFor:FieldComments.Text
        //ExSummary:Shows how to use the COMMENTS field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set a value for the document's "Comments" built-in property.
        doc->get_BuiltInDocumentProperties()->set_Comments(u"My comment.");

        // Create a COMMENTS field to display the value of that built-in property.
        auto field = System::DynamicCast<FieldComments>(builder->InsertField(FieldType::FieldComments, true));
        field->Update();

        ASSERT_EQ(u" COMMENTS ", field->GetFieldCode());
        ASSERT_EQ(u"My comment.", field->get_Result());

        // If we give the COMMENTS field's Text property value and update it, the field will
        // overwrite the current value of the "Comments" built-in property with the value of its Text property,
        // and then display the new value.
        field->set_Text(u"My overriding comment.");
        field->Update();

        ASSERT_EQ(u" COMMENTS  \"My overriding comment.\"", field->GetFieldCode());
        ASSERT_EQ(u"My overriding comment.", field->get_Result());

        doc->Save(ArtifactsDir + u"Field.COMMENTS.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.COMMENTS.docx");

        ASSERT_EQ(u"My overriding comment.", doc->get_BuiltInDocumentProperties()->get_Comments());

        field = System::DynamicCast<FieldComments>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldComments, u" COMMENTS  \"My overriding comment.\"", u"My overriding comment.", field);
        ASSERT_EQ(u"My overriding comment.", field->get_Text());
    }

    void FieldFileSize_()
    {
        //ExStart
        //ExFor:FieldFileSize
        //ExFor:FieldFileSize.IsInKilobytes
        //ExFor:FieldFileSize.IsInMegabytes
        //ExSummary:Shows how to display the file size of a document with a FILESIZE field.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        ASSERT_EQ(18105, doc->get_BuiltInDocumentProperties()->get_Bytes());

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveToDocumentEnd();
        builder->InsertParagraph();

        // Below are three different units of measure
        // with which FILESIZE fields can display the document's file size.
        // 1 -  Bytes:
        auto field = System::DynamicCast<FieldFileSize>(builder->InsertField(FieldType::FieldFileSize, true));
        field->Update();

        ASSERT_EQ(u" FILESIZE ", field->GetFieldCode());
        ASSERT_EQ(u"18105", field->get_Result());

        // 2 -  Kilobytes:
        builder->InsertParagraph();
        field = System::DynamicCast<FieldFileSize>(builder->InsertField(FieldType::FieldFileSize, true));
        field->set_IsInKilobytes(true);
        field->Update();

        ASSERT_EQ(u" FILESIZE  \\k", field->GetFieldCode());
        ASSERT_EQ(u"18", field->get_Result());

        // 3 -  Megabytes:
        builder->InsertParagraph();
        field = System::DynamicCast<FieldFileSize>(builder->InsertField(FieldType::FieldFileSize, true));
        field->set_IsInMegabytes(true);
        field->Update();

        ASSERT_EQ(u" FILESIZE  \\m", field->GetFieldCode());
        ASSERT_EQ(u"0", field->get_Result());

        // To update the values of these fields while editing in Microsoft Word,
        // we must first save the changes, and then manually update these fields.
        doc->Save(ArtifactsDir + u"Field.FILESIZE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.FILESIZE.docx");

        field = System::DynamicCast<FieldFileSize>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldFileSize, u" FILESIZE ", u"18105", field);

        // These fields will need to be updated to produce an accurate result.
        doc->UpdateFields();

        field = System::DynamicCast<FieldFileSize>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldFileSize, u" FILESIZE  \\k", u"13", field);
        ASSERT_TRUE(field->get_IsInKilobytes());

        field = System::DynamicCast<FieldFileSize>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldFileSize, u" FILESIZE  \\m", u"0", field);
        ASSERT_TRUE(field->get_IsInMegabytes());
    }

    void FieldGoToButton_()
    {
        //ExStart
        //ExFor:FieldGoToButton
        //ExFor:FieldGoToButton.DisplayText
        //ExFor:FieldGoToButton.Location
        //ExSummary:Shows to insert a GOTOBUTTON field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add a GOTOBUTTON field. When we double-click this field in Microsoft Word,
        // it will take the text cursor to the bookmark whose name the Location property references.
        auto field = System::DynamicCast<FieldGoToButton>(builder->InsertField(FieldType::FieldGoToButton, true));
        field->set_DisplayText(u"My Button");
        field->set_Location(u"MyBookmark");

        ASSERT_EQ(u" GOTOBUTTON  MyBookmark My Button", field->GetFieldCode());

        // Insert a valid bookmark for the field to reference.
        builder->InsertBreak(BreakType::PageBreak);
        builder->StartBookmark(field->get_Location());
        builder->Writeln(u"Bookmark text contents.");
        builder->EndBookmark(field->get_Location());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.GOTOBUTTON.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.GOTOBUTTON.docx");
        field = System::DynamicCast<FieldGoToButton>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldGoToButton, u" GOTOBUTTON  MyBookmark My Button", String::Empty, field);
        ASSERT_EQ(u"My Button", field->get_DisplayText());
        ASSERT_EQ(u"MyBookmark", field->get_Location());
    }

    //ExStart
    //ExFor:FieldFillIn
    //ExFor:FieldFillIn.DefaultResponse
    //ExFor:FieldFillIn.PromptOnceOnMailMerge
    //ExFor:FieldFillIn.PromptText
    //ExSummary:Shows how to use the FILLIN field to prompt the user for a response.
    void FieldFillIn_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a FILLIN field. When we manually update this field in Microsoft Word,
        // it will prompt us to enter a response. The field will then display the response as text.
        auto field = System::DynamicCast<FieldFillIn>(builder->InsertField(FieldType::FieldFillIn, true));
        field->set_PromptText(u"Please enter a response:");
        field->set_DefaultResponse(u"A default response.");

        // We can also use these fields to ask the user for a unique response for each page
        // created during a mail merge done using Microsoft Word.
        field->set_PromptOnceOnMailMerge(true);

        ASSERT_EQ(u" FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o", field->GetFieldCode());

        auto mergeField = System::DynamicCast<FieldMergeField>(builder->InsertField(FieldType::FieldMergeField, true));
        mergeField->set_FieldName(u"MergeField");

        // If we perform a mail merge programmatically, we can use a custom prompt respondent
        // to automatically edit responses for FILLIN fields that the mail merge encounters.
        doc->get_FieldOptions()->set_UserPromptRespondent(MakeObject<ExField::PromptRespondent>());
        doc->get_MailMerge()->Execute(MakeArray<String>({u"MergeField"}), MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(u"")}));

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.FILLIN.docx");
        TestFieldFillIn(MakeObject<Document>(ArtifactsDir + u"Field.FILLIN.docx"));
        //ExSkip
    }

    /// <summary>
    /// Prepends a line to the default response of every FILLIN field during a mail merge.
    /// </summary>
    class PromptRespondent : public IFieldUserPromptRespondent
    {
    public:
        String Respond(String promptText, String defaultResponse) override
        {
            return String(u"Response modified by PromptRespondent. ") + defaultResponse;
        }
    };
    //ExEnd

    void TestFieldFillIn(SharedPtr<Document> doc)
    {
        doc = DocumentHelper::SaveOpen(doc);

        ASSERT_EQ(1, doc->get_Range()->get_Fields()->get_Count());

        auto field = System::DynamicCast<FieldFillIn>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldFillIn, u" FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o",
                              u"Response modified by PromptRespondent. A default response.", field);
        ASSERT_EQ(u"Please enter a response:", field->get_PromptText());
        ASSERT_EQ(u"A default response.", field->get_DefaultResponse());
        ASSERT_TRUE(field->get_PromptOnceOnMailMerge());
    }

    void FieldInfo_()
    {
        //ExStart
        //ExFor:FieldInfo
        //ExFor:FieldInfo.InfoType
        //ExFor:FieldInfo.NewValue
        //ExSummary:Shows how to work with INFO fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set a value for the "Comments" built-in property and then insert an INFO field to display that property's value.
        doc->get_BuiltInDocumentProperties()->set_Comments(u"My comment");
        auto field = System::DynamicCast<FieldInfo>(builder->InsertField(FieldType::FieldInfo, true));
        field->set_InfoType(u"Comments");
        field->Update();

        ASSERT_EQ(u" INFO  Comments", field->GetFieldCode());
        ASSERT_EQ(u"My comment", field->get_Result());

        builder->Writeln();

        // Setting a value for the field's NewValue property and updating
        // the field will also overwrite the corresponding built-in property with the new value.
        field = System::DynamicCast<FieldInfo>(builder->InsertField(FieldType::FieldInfo, true));
        field->set_InfoType(u"Comments");
        field->set_NewValue(u"New comment");
        field->Update();

        ASSERT_EQ(u" INFO  Comments \"New comment\"", field->GetFieldCode());
        ASSERT_EQ(u"New comment", field->get_Result());
        ASSERT_EQ(u"New comment", doc->get_BuiltInDocumentProperties()->get_Comments());

        doc->Save(ArtifactsDir + u"Field.INFO.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.INFO.docx");

        ASSERT_EQ(u"New comment", doc->get_BuiltInDocumentProperties()->get_Comments());

        field = System::DynamicCast<FieldInfo>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldInfo, u" INFO  Comments", u"My comment", field);
        ASSERT_EQ(u"Comments", field->get_InfoType());

        field = System::DynamicCast<FieldInfo>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldInfo, u" INFO  Comments \"New comment\"", u"New comment", field);
        ASSERT_EQ(u"Comments", field->get_InfoType());
        ASSERT_EQ(u"New comment", field->get_NewValue());
    }

    void FieldMacroButton_()
    {
        //ExStart
        //ExFor:Document.HasMacros
        //ExFor:FieldMacroButton
        //ExFor:FieldMacroButton.DisplayText
        //ExFor:FieldMacroButton.MacroName
        //ExSummary:Shows how to use MACROBUTTON fields to allow us to run a document's macros by clicking.
        auto doc = MakeObject<Document>(MyDir + u"Macro.docm");
        auto builder = MakeObject<DocumentBuilder>(doc);

        ASSERT_TRUE(doc->get_HasMacros());

        // Insert a MACROBUTTON field, and reference one of the document's macros by name in the MacroName property.
        auto field = System::DynamicCast<FieldMacroButton>(builder->InsertField(FieldType::FieldMacroButton, true));
        field->set_MacroName(u"MyMacro");
        field->set_DisplayText(String(u"Double click to run macro: ") + field->get_MacroName());

        ASSERT_EQ(u" MACROBUTTON  MyMacro Double click to run macro: MyMacro", field->GetFieldCode());

        // Use the property to reference "ViewZoom200", a macro that ships with Microsoft Word.
        // We can find all other macros via View -> Macros (dropdown) -> View Macros.
        // In that menu, select "Word Commands" from the "Macros in:" drop down.
        // If our document contains a custom macro with the same name as a stock macro,
        // our macro will be the one that the MACROBUTTON field runs.
        builder->InsertParagraph();
        field = System::DynamicCast<FieldMacroButton>(builder->InsertField(FieldType::FieldMacroButton, true));
        field->set_MacroName(u"ViewZoom200");
        field->set_DisplayText(String(u"Run ") + field->get_MacroName());

        ASSERT_EQ(u" MACROBUTTON  ViewZoom200 Run ViewZoom200", field->GetFieldCode());

        // Save the document as a macro-enabled document type.
        doc->Save(ArtifactsDir + u"Field.MACROBUTTON.docm");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.MACROBUTTON.docm");

        field = System::DynamicCast<FieldMacroButton>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldMacroButton, u" MACROBUTTON  MyMacro Double click to run macro: MyMacro", String::Empty, field);
        ASSERT_EQ(u"MyMacro", field->get_MacroName());
        ASSERT_EQ(u"Double click to run macro: MyMacro", field->get_DisplayText());

        field = System::DynamicCast<FieldMacroButton>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldMacroButton, u" MACROBUTTON  ViewZoom200 Run ViewZoom200", String::Empty, field);
        ASSERT_EQ(u"ViewZoom200", field->get_MacroName());
        ASSERT_EQ(u"Run ViewZoom200", field->get_DisplayText());
    }

    void FieldKeywords_()
    {
        //ExStart
        //ExFor:FieldKeywords
        //ExFor:FieldKeywords.Text
        //ExSummary:Shows to insert a KEYWORDS field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add some keywords, also referred to as "tags" in File Explorer.
        doc->get_BuiltInDocumentProperties()->set_Keywords(u"Keyword1, Keyword2");

        // The KEYWORDS field displays the value of this property.
        auto field = System::DynamicCast<FieldKeywords>(builder->InsertField(FieldType::FieldKeyword, true));
        field->Update();

        ASSERT_EQ(u" KEYWORDS ", field->GetFieldCode());
        ASSERT_EQ(u"Keyword1, Keyword2", field->get_Result());

        // Setting a value for the field's Text property,
        // and then updating the field will also overwrite the corresponding built-in property with the new value.
        field->set_Text(u"OverridingKeyword");
        field->Update();

        ASSERT_EQ(u" KEYWORDS  OverridingKeyword", field->GetFieldCode());
        ASSERT_EQ(u"OverridingKeyword", field->get_Result());
        ASSERT_EQ(u"OverridingKeyword", doc->get_BuiltInDocumentProperties()->get_Keywords());

        doc->Save(ArtifactsDir + u"Field.KEYWORDS.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.KEYWORDS.docx");

        ASSERT_EQ(u"OverridingKeyword", doc->get_BuiltInDocumentProperties()->get_Keywords());

        field = System::DynamicCast<FieldKeywords>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldKeyword, u" KEYWORDS  OverridingKeyword", u"OverridingKeyword", field);
        ASSERT_EQ(u"OverridingKeyword", field->get_Text());
    }

    void FieldNum()
    {
        //ExStart
        //ExFor:FieldPage
        //ExFor:FieldNumChars
        //ExFor:FieldNumPages
        //ExFor:FieldNumWords
        //ExSummary:Shows how to use NUMCHARS, NUMWORDS, NUMPAGES and PAGE fields to track the size of our documents.
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);

        // Below are three types of fields that we can use to track the size of our documents.
        // 1 -  Track the character count with a NUMCHARS field:
        auto fieldNumChars = System::DynamicCast<FieldNumChars>(builder->InsertField(FieldType::FieldNumChars, true));
        builder->Writeln(u" characters");

        // 2 -  Track the word count with a NUMWORDS field:
        auto fieldNumWords = System::DynamicCast<FieldNumWords>(builder->InsertField(FieldType::FieldNumWords, true));
        builder->Writeln(u" words");

        // 3 -  Use both PAGE and NUMPAGES fields to display what page the field is on,
        // and the total number of pages in the document:
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);
        builder->Write(u"Page ");
        auto fieldPage = System::DynamicCast<FieldPage>(builder->InsertField(FieldType::FieldPage, true));
        builder->Write(u" of ");
        auto fieldNumPages = System::DynamicCast<FieldNumPages>(builder->InsertField(FieldType::FieldNumPages, true));

        ASSERT_EQ(u" NUMCHARS ", fieldNumChars->GetFieldCode());
        ASSERT_EQ(u" NUMWORDS ", fieldNumWords->GetFieldCode());
        ASSERT_EQ(u" NUMPAGES ", fieldNumPages->GetFieldCode());
        ASSERT_EQ(u" PAGE ", fieldPage->GetFieldCode());

        // These fields will not maintain accurate values in real time
        // while we edit the document programmatically using Aspose.Words, or in Microsoft Word.
        // We need to update them every we need to see an up-to-date value.
        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.NUMCHARS.NUMWORDS.NUMPAGES.PAGE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.NUMCHARS.NUMWORDS.NUMPAGES.PAGE.docx");

        TestUtil::VerifyField(FieldType::FieldNumChars, u" NUMCHARS ", u"6009", doc->get_Range()->get_Fields()->idx_get(0));
        TestUtil::VerifyField(FieldType::FieldNumWords, u" NUMWORDS ", u"1054", doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldPage, u" PAGE ", u"6", doc->get_Range()->get_Fields()->idx_get(2));
        TestUtil::VerifyField(FieldType::FieldNumPages, u" NUMPAGES ", u"6", doc->get_Range()->get_Fields()->idx_get(3));
    }

    void FieldPrint_()
    {
        //ExStart
        //ExFor:FieldPrint
        //ExFor:FieldPrint.PostScriptGroup
        //ExFor:FieldPrint.PrinterInstructions
        //ExSummary:Shows to insert a PRINT field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"My paragraph");

        // The PRINT field can send instructions to the printer.
        auto field = System::DynamicCast<FieldPrint>(builder->InsertField(FieldType::FieldPrint, true));

        // Set the area for the printer to perform instructions over.
        // In this case, it will be the paragraph that contains our PRINT field.
        field->set_PostScriptGroup(u"para");

        // When we use a printer that supports PostScript to print our document,
        // this command will turn the entire area that we specified in "field.PostScriptGroup" white.
        field->set_PrinterInstructions(u"erasepage");

        ASSERT_EQ(u" PRINT  erasepage \\p para", field->GetFieldCode());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.PRINT.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.PRINT.docx");

        field = System::DynamicCast<FieldPrint>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldPrint, u" PRINT  erasepage \\p para", String::Empty, field);
        ASSERT_EQ(u"para", field->get_PostScriptGroup());
        ASSERT_EQ(u"erasepage", field->get_PrinterInstructions());
    }

    void FieldPrintDate_()
    {
        //ExStart
        //ExFor:FieldPrintDate
        //ExFor:FieldPrintDate.UseLunarCalendar
        //ExFor:FieldPrintDate.UseSakaEraCalendar
        //ExFor:FieldPrintDate.UseUmAlQuraCalendar
        //ExSummary:Shows read PRINTDATE fields.
        auto doc = MakeObject<Document>(MyDir + u"Field sample - PRINTDATE.docx");

        // When a document is printed by a printer or printed as a PDF (but not exported to PDF),
        // PRINTDATE fields will display the print operation's date/time.
        // If no printing has taken place, these fields will display "0/0/0000".
        auto field = System::DynamicCast<FieldPrintDate>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(u"3/25/2020 12:00:00 AM", field->get_Result());
        ASSERT_EQ(u" PRINTDATE ", field->GetFieldCode());

        // Below are three different calendar types according to which the PRINTDATE field
        // can display the date and time of the last printing operation.
        // 1 -  Islamic Lunar Calendar:
        field = System::DynamicCast<FieldPrintDate>(doc->get_Range()->get_Fields()->idx_get(1));

        ASSERT_TRUE(field->get_UseLunarCalendar());
        ASSERT_EQ(u"8/1/1441 12:00:00 AM", field->get_Result());
        ASSERT_EQ(u" PRINTDATE  \\h", field->GetFieldCode());

        field = System::DynamicCast<FieldPrintDate>(doc->get_Range()->get_Fields()->idx_get(2));

        // 2 -  Umm al-Qura calendar:
        ASSERT_TRUE(field->get_UseUmAlQuraCalendar());
        ASSERT_EQ(u"8/1/1441 12:00:00 AM", field->get_Result());
        ASSERT_EQ(u" PRINTDATE  \\u", field->GetFieldCode());

        field = System::DynamicCast<FieldPrintDate>(doc->get_Range()->get_Fields()->idx_get(3));

        // 3 -  Indian National Calendar:
        ASSERT_TRUE(field->get_UseSakaEraCalendar());
        ASSERT_EQ(u"1/5/1942 12:00:00 AM", field->get_Result());
        ASSERT_EQ(u" PRINTDATE  \\s", field->GetFieldCode());
        //ExEnd
    }

    void FieldQuote_()
    {
        //ExStart
        //ExFor:FieldQuote
        //ExFor:FieldQuote.Text
        //ExFor:Document.UpdateFields
        //ExSummary:Shows to use the QUOTE field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a QUOTE field, which will display the value of its Text property.
        auto field = System::DynamicCast<FieldQuote>(builder->InsertField(FieldType::FieldQuote, true));
        field->set_Text(u"\"Quoted text\"");

        ASSERT_EQ(u" QUOTE  \"\\\"Quoted text\\\"\"", field->GetFieldCode());

        // Insert a QUOTE field and nest a DATE field inside it.
        // DATE fields update their value to the current date every time we open the document using Microsoft Word.
        // Nesting the DATE field inside the QUOTE field like this will freeze its value
        // to the date when we created the document.
        builder->Write(u"\nDocument creation date: ");
        field = System::DynamicCast<FieldQuote>(builder->InsertField(FieldType::FieldQuote, true));
        builder->MoveTo(field->get_Separator());
        builder->InsertField(FieldType::FieldDate, true);

        ASSERT_EQ(String(u" QUOTE \u0013 DATE \u0014") + System::DateTime::get_Now().get_Date().ToShortDateString() + u"\u0015", field->GetFieldCode());

        // Update all the fields to display their correct results.
        doc->UpdateFields();

        ASSERT_EQ(u"\"Quoted text\"", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());

        doc->Save(ArtifactsDir + u"Field.QUOTE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.QUOTE.docx");

        TestUtil::VerifyField(FieldType::FieldQuote, u" QUOTE  \"\\\"Quoted text\\\"\"", u"\"Quoted text\"", doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldQuote,
                              String(u" QUOTE \u0013 DATE \u0014") + System::DateTime::get_Now().get_Date().ToShortDateString() + u"\u0015",
                              System::DateTime::get_Now().get_Date().ToShortDateString(), doc->get_Range()->get_Fields()->idx_get(1));
    }

    //ExStart
    //ExFor:FieldNoteRef
    //ExFor:FieldNoteRef.BookmarkName
    //ExFor:FieldNoteRef.InsertHyperlink
    //ExFor:FieldNoteRef.InsertReferenceMark
    //ExFor:FieldNoteRef.InsertRelativePosition
    //ExSummary:Shows to insert NOTEREF fields, and modify their appearance.
    void FieldNoteRef_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a bookmark with a footnote that the NOTEREF field will reference.
        InsertBookmarkWithFootnote(builder, u"MyBookmark1", u"Contents of MyBookmark1", u"Footnote from MyBookmark1");

        // This NOTEREF field will display the number of the footnote inside the referenced bookmark.
        // Setting the InsertHyperlink property lets us jump to the bookmark by Ctrl + clicking the field in Microsoft Word.
        ASSERT_EQ(u" NOTEREF  MyBookmark2 \\h",
                  InsertFieldNoteRef(builder, u"MyBookmark2", true, false, false, u"Hyperlink to Bookmark2, with footnote number ")->GetFieldCode());

        // When using the \p flag, after the footnote number, the field also displays the bookmark's position relative to the field.
        // Bookmark1 is above this field and contains footnote number 1, so the result will be "1 above" on update.
        ASSERT_EQ(u" NOTEREF  MyBookmark1 \\h \\p",
                  InsertFieldNoteRef(builder, u"MyBookmark1", true, true, false, u"Bookmark1, with footnote number ")->GetFieldCode());

        // Bookmark2 is below this field and contains footnote number 2, so the field will display "2 below".
        // The \f flag makes the number 2 appear in the same format as the footnote number label in the actual text.
        ASSERT_EQ(u" NOTEREF  MyBookmark2 \\h \\p \\f",
                  InsertFieldNoteRef(builder, u"MyBookmark2", true, true, true, u"Bookmark2, with footnote number ")->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);
        InsertBookmarkWithFootnote(builder, u"MyBookmark2", u"Contents of MyBookmark2", u"Footnote from MyBookmark2");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.NOTEREF.docx");
        TestNoteRef(MakeObject<Document>(ArtifactsDir + u"Field.NOTEREF.docx"));
        //ExSkip
    }

    /// <summary>
    /// Uses a document builder to insert a NOTEREF field with specified properties.
    /// </summary>
    static SharedPtr<FieldNoteRef> InsertFieldNoteRef(SharedPtr<DocumentBuilder> builder, String bookmarkName, bool insertHyperlink,
                                                      bool insertRelativePosition, bool insertReferenceMark, String textBefore)
    {
        builder->Write(textBefore);

        auto field = System::DynamicCast<FieldNoteRef>(builder->InsertField(FieldType::FieldNoteRef, true));
        field->set_BookmarkName(bookmarkName);
        field->set_InsertHyperlink(insertHyperlink);
        field->set_InsertRelativePosition(insertRelativePosition);
        field->set_InsertReferenceMark(insertReferenceMark);
        builder->Writeln();

        return field;
    }

    /// <summary>
    /// Uses a document builder to insert a named bookmark with a footnote at the end.
    /// </summary>
    static void InsertBookmarkWithFootnote(SharedPtr<DocumentBuilder> builder, String bookmarkName, String bookmarkText, String footnoteText)
    {
        builder->StartBookmark(bookmarkName);
        builder->Write(bookmarkText);
        builder->InsertFootnote(FootnoteType::Footnote, footnoteText);
        builder->EndBookmark(bookmarkName);
        builder->Writeln();
    }
    //ExEnd

    void TestNoteRef(SharedPtr<Document> doc)
    {
        auto field = System::DynamicCast<FieldNoteRef>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldNoteRef, u" NOTEREF  MyBookmark2 \\h", u"2", field);
        ASSERT_EQ(u"MyBookmark2", field->get_BookmarkName());
        ASSERT_TRUE(field->get_InsertHyperlink());
        ASSERT_FALSE(field->get_InsertRelativePosition());
        ASSERT_FALSE(field->get_InsertReferenceMark());

        field = System::DynamicCast<FieldNoteRef>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldNoteRef, u" NOTEREF  MyBookmark1 \\h \\p", u"1 above", field);
        ASSERT_EQ(u"MyBookmark1", field->get_BookmarkName());
        ASSERT_TRUE(field->get_InsertHyperlink());
        ASSERT_TRUE(field->get_InsertRelativePosition());
        ASSERT_FALSE(field->get_InsertReferenceMark());

        field = System::DynamicCast<FieldNoteRef>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldNoteRef, u" NOTEREF  MyBookmark2 \\h \\p \\f", u"2 below", field);
        ASSERT_EQ(u"MyBookmark2", field->get_BookmarkName());
        ASSERT_TRUE(field->get_InsertHyperlink());
        ASSERT_TRUE(field->get_InsertRelativePosition());
        ASSERT_TRUE(field->get_InsertReferenceMark());
    }

    void FootnoteRef()
    {
        //ExStart
        //ExFor:FieldFootnoteRef
        //ExSummary:Shows how to cross-reference footnotes with the FOOTNOTEREF field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartBookmark(u"CrossRefBookmark");
        builder->Write(u"Hello world!");
        builder->InsertFootnote(FootnoteType::Footnote, u"Cross referenced footnote.");
        builder->EndBookmark(u"CrossRefBookmark");
        builder->InsertParagraph();

        // Insert a FOOTNOTEREF field, which lets us reference a footnote more than once while re-using the same footnote marker.
        builder->Write(u"CrossReference: ");
        auto field = System::DynamicCast<FieldFootnoteRef>(builder->InsertField(FieldType::FieldFootnoteRef, true));

        // Reference the bookmark that we have created with the FOOTNOTEREF field. That bookmark contains a footnote marker
        // belonging to the footnote we inserted. The field will display that footnote marker.
        builder->MoveTo(field->get_Separator());
        builder->Write(u"CrossRefBookmark");

        ASSERT_EQ(u" FOOTNOTEREF CrossRefBookmark", field->GetFieldCode());

        doc->UpdateFields();

        // This field works only in older versions of Microsoft Word.
        doc->Save(ArtifactsDir + u"Field.FOOTNOTEREF.doc");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.FOOTNOTEREF.doc");
        field = System::DynamicCast<FieldFootnoteRef>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldFootnoteRef, u" FOOTNOTEREF CrossRefBookmark", u"1", field);
        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"Cross referenced footnote.",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
    }

    //ExStart
    //ExFor:FieldPageRef
    //ExFor:FieldPageRef.BookmarkName
    //ExFor:FieldPageRef.InsertHyperlink
    //ExFor:FieldPageRef.InsertRelativePosition
    //ExSummary:Shows to insert PAGEREF fields to display the relative location of bookmarks.
    void FieldPageRef_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        InsertAndNameBookmark(builder, u"MyBookmark1");

        // Insert a PAGEREF field that displays what page a bookmark is on.
        // Set the InsertHyperlink flag to make the field also function as a clickable link to the bookmark.
        ASSERT_EQ(u" PAGEREF  MyBookmark3 \\h", InsertFieldPageRef(builder, u"MyBookmark3", true, false, u"Hyperlink to Bookmark3, on page: ")->GetFieldCode());

        // We can use the \p flag to get the PAGEREF field to display
        // the bookmark's position relative to the position of the field.
        // Bookmark1 is on the same page and above this field, so this field's displayed result will be "above".
        ASSERT_EQ(u" PAGEREF  MyBookmark1 \\h \\p", InsertFieldPageRef(builder, u"MyBookmark1", true, true, u"Bookmark1 is ")->GetFieldCode());

        // Bookmark2 will be on the same page and below this field, so this field's displayed result will be "below".
        ASSERT_EQ(u" PAGEREF  MyBookmark2 \\h \\p", InsertFieldPageRef(builder, u"MyBookmark2", true, true, u"Bookmark2 is ")->GetFieldCode());

        // Bookmark3 will be on a different page, so the field will display "on page 2".
        ASSERT_EQ(u" PAGEREF  MyBookmark3 \\h \\p", InsertFieldPageRef(builder, u"MyBookmark3", true, true, u"Bookmark3 is ")->GetFieldCode());

        InsertAndNameBookmark(builder, u"MyBookmark2");
        builder->InsertBreak(BreakType::PageBreak);
        InsertAndNameBookmark(builder, u"MyBookmark3");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.PAGEREF.docx");
        TestPageRef(MakeObject<Document>(ArtifactsDir + u"Field.PAGEREF.docx"));
        //ExSkip
    }

    /// <summary>
    /// Uses a document builder to insert a PAGEREF field and sets its properties.
    /// </summary>
    static SharedPtr<FieldPageRef> InsertFieldPageRef(SharedPtr<DocumentBuilder> builder, String bookmarkName, bool insertHyperlink,
                                                      bool insertRelativePosition, String textBefore)
    {
        builder->Write(textBefore);

        auto field = System::DynamicCast<FieldPageRef>(builder->InsertField(FieldType::FieldPageRef, true));
        field->set_BookmarkName(bookmarkName);
        field->set_InsertHyperlink(insertHyperlink);
        field->set_InsertRelativePosition(insertRelativePosition);
        builder->Writeln();

        return field;
    }

    /// <summary>
    /// Uses a document builder to insert a named bookmark.
    /// </summary>
    static void InsertAndNameBookmark(SharedPtr<DocumentBuilder> builder, String bookmarkName)
    {
        builder->StartBookmark(bookmarkName);
        builder->Writeln(String::Format(u"Contents of bookmark \"{0}\".", bookmarkName));
        builder->EndBookmark(bookmarkName);
    }
    //ExEnd

    void TestPageRef(SharedPtr<Document> doc)
    {
        auto field = System::DynamicCast<FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF  MyBookmark3 \\h", u"2", field);
        ASSERT_EQ(u"MyBookmark3", field->get_BookmarkName());
        ASSERT_TRUE(field->get_InsertHyperlink());
        ASSERT_FALSE(field->get_InsertRelativePosition());

        field = System::DynamicCast<FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF  MyBookmark1 \\h \\p", u"above", field);
        ASSERT_EQ(u"MyBookmark1", field->get_BookmarkName());
        ASSERT_TRUE(field->get_InsertHyperlink());
        ASSERT_TRUE(field->get_InsertRelativePosition());

        field = System::DynamicCast<FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF  MyBookmark2 \\h \\p", u"below", field);
        ASSERT_EQ(u"MyBookmark2", field->get_BookmarkName());
        ASSERT_TRUE(field->get_InsertHyperlink());
        ASSERT_TRUE(field->get_InsertRelativePosition());

        field = System::DynamicCast<FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF  MyBookmark3 \\h \\p", u"on page 2", field);
        ASSERT_EQ(u"MyBookmark3", field->get_BookmarkName());
        ASSERT_TRUE(field->get_InsertHyperlink());
        ASSERT_TRUE(field->get_InsertRelativePosition());
    }

    //ExStart
    //ExFor:FieldRef
    //ExFor:FieldRef.BookmarkName
    //ExFor:FieldRef.IncludeNoteOrComment
    //ExFor:FieldRef.InsertHyperlink
    //ExFor:FieldRef.InsertParagraphNumber
    //ExFor:FieldRef.InsertParagraphNumberInFullContext
    //ExFor:FieldRef.InsertParagraphNumberInRelativeContext
    //ExFor:FieldRef.InsertRelativePosition
    //ExFor:FieldRef.NumberSeparator
    //ExFor:FieldRef.SuppressNonDelimiters
    //ExSummary:Shows how to insert REF fields to reference bookmarks.
    void FieldRef_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartBookmark(u"MyBookmark");
        builder->InsertFootnote(FootnoteType::Footnote, u"MyBookmark footnote #1");
        builder->Write(u"Text that will appear in REF field");
        builder->InsertFootnote(FootnoteType::Footnote, u"MyBookmark footnote #2");
        builder->EndBookmark(u"MyBookmark");
        builder->MoveToDocumentStart();

        // We will apply a custom list format, where the amount of angle brackets indicates the list level we are currently at.
        builder->get_ListFormat()->ApplyNumberDefault();
        builder->get_ListFormat()->get_ListLevel()->set_NumberFormat(u"> \x0000");

        // Insert a REF field that will contain the text within our bookmark, act as a hyperlink, and clone the bookmark's footnotes.
        SharedPtr<FieldRef> field = InsertFieldRef(builder, u"MyBookmark", u"", u"\n");
        field->set_IncludeNoteOrComment(true);
        field->set_InsertHyperlink(true);

        ASSERT_EQ(u" REF  MyBookmark \\f \\h", field->GetFieldCode());

        // Insert a REF field, and display whether the referenced bookmark is above or below it.
        field = InsertFieldRef(builder, u"MyBookmark", u"The referenced paragraph is ", u" this field.\n");
        field->set_InsertRelativePosition(true);

        ASSERT_EQ(u" REF  MyBookmark \\p", field->GetFieldCode());

        // Display the list number of the bookmark as it appears in the document.
        field = InsertFieldRef(builder, u"MyBookmark", u"The bookmark's paragraph number is ", u"\n");
        field->set_InsertParagraphNumber(true);

        ASSERT_EQ(u" REF  MyBookmark \\n", field->GetFieldCode());

        // Display the bookmark's list number, but with non-delimiter characters, such as the angle brackets, omitted.
        field = InsertFieldRef(builder, u"MyBookmark", u"The bookmark's paragraph number, non-delimiters suppressed, is ", u"\n");
        field->set_InsertParagraphNumber(true);
        field->set_SuppressNonDelimiters(true);

        ASSERT_EQ(u" REF  MyBookmark \\n \\t", field->GetFieldCode());

        // Move down one list level.
        builder->get_ListFormat()->set_ListLevelNumber(builder->get_ListFormat()->get_ListLevelNumber() + 1);
        builder->get_ListFormat()->get_ListLevel()->set_NumberFormat(u">> \x0001");

        // Display the list number of the bookmark and the numbers of all the list levels above it.
        field = InsertFieldRef(builder, u"MyBookmark", u"The bookmark's full context paragraph number is ", u"\n");
        field->set_InsertParagraphNumberInFullContext(true);

        ASSERT_EQ(u" REF  MyBookmark \\w", field->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);

        // Display the list level numbers between this REF field, and the bookmark that it is referencing.
        field = InsertFieldRef(builder, u"MyBookmark", u"The bookmark's relative paragraph number is ", u"\n");
        field->set_InsertParagraphNumberInRelativeContext(true);

        ASSERT_EQ(u" REF  MyBookmark \\r", field->GetFieldCode());

        // At the end of the document, the bookmark will show up as a list item here.
        builder->Writeln(u"List level above bookmark");
        builder->get_ListFormat()->set_ListLevelNumber(builder->get_ListFormat()->get_ListLevelNumber() + 1);
        builder->get_ListFormat()->get_ListLevel()->set_NumberFormat(u">>> \x0002");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.REF.docx");
        TestFieldRef(MakeObject<Document>(ArtifactsDir + u"Field.REF.docx"));
        //ExSkip
    }

    /// <summary>
    /// Get the document builder to insert a REF field, reference a bookmark with it, and add text before and after it.
    /// </summary>
    static SharedPtr<FieldRef> InsertFieldRef(SharedPtr<DocumentBuilder> builder, String bookmarkName, String textBefore, String textAfter)
    {
        builder->Write(textBefore);
        auto field = System::DynamicCast<FieldRef>(builder->InsertField(FieldType::FieldRef, true));
        field->set_BookmarkName(bookmarkName);
        builder->Write(textAfter);
        return field;
    }
    //ExEnd

    void TestFieldRef(SharedPtr<Document> doc)
    {
        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"MyBookmark footnote #1",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
        TestUtil::VerifyFootnote(FootnoteType::Footnote, true, String::Empty, u"MyBookmark footnote #2",
                                 System::DynamicCast<Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));

        auto field = System::DynamicCast<FieldRef>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark \\f \\h",
                              String(u"\u0002 MyBookmark footnote #1\r") + u"Text that will appear in REF field\u0002 MyBookmark footnote #2\r", field);
        ASSERT_EQ(u"MyBookmark", field->get_BookmarkName());
        ASSERT_TRUE(field->get_IncludeNoteOrComment());
        ASSERT_TRUE(field->get_InsertHyperlink());

        field = System::DynamicCast<FieldRef>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark \\p", u"below", field);
        ASSERT_EQ(u"MyBookmark", field->get_BookmarkName());
        ASSERT_TRUE(field->get_InsertRelativePosition());

        field = System::DynamicCast<FieldRef>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark \\n", u">>> i", field);
        ASSERT_EQ(u"MyBookmark", field->get_BookmarkName());
        ASSERT_TRUE(field->get_InsertParagraphNumber());
        ASSERT_EQ(u" REF  MyBookmark \\n", field->GetFieldCode());
        ASSERT_EQ(u">>> i", field->get_Result());

        field = System::DynamicCast<FieldRef>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark \\n \\t", u"i", field);
        ASSERT_EQ(u"MyBookmark", field->get_BookmarkName());
        ASSERT_TRUE(field->get_InsertParagraphNumber());
        ASSERT_TRUE(field->get_SuppressNonDelimiters());

        field = System::DynamicCast<FieldRef>(doc->get_Range()->get_Fields()->idx_get(4));

        TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark \\w", u"> 4>> c>>> i", field);
        ASSERT_EQ(u"MyBookmark", field->get_BookmarkName());
        ASSERT_TRUE(field->get_InsertParagraphNumberInFullContext());

        field = System::DynamicCast<FieldRef>(doc->get_Range()->get_Fields()->idx_get(5));

        TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark \\r", u">> c>>> i", field);
        ASSERT_EQ(u"MyBookmark", field->get_BookmarkName());
        ASSERT_TRUE(field->get_InsertParagraphNumberInRelativeContext());
    }

    void FieldRD_()
    {
        //ExStart
        //ExFor:FieldRD
        //ExFor:FieldRD.FileName
        //ExFor:FieldRD.IsPathRelative
        //ExSummary:Shows to use the RD field to create a table of contents entries from headings in other documents.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to insert a table of contents,
        // and then add one entry for the table of contents on the following page.
        builder->InsertField(FieldType::FieldTOC, true);
        builder->InsertBreak(BreakType::PageBreak);
        builder->get_CurrentParagraph()->get_ParagraphFormat()->set_StyleName(u"Heading 1");
        builder->Writeln(u"TOC entry from within this document");

        // Insert an RD field, which references another local file system document in its FileName property.
        // The TOC will also now accept all headings from the referenced document as entries for its table.
        auto field = System::DynamicCast<FieldRD>(builder->InsertField(FieldType::FieldRefDoc, true));
        field->set_FileName(u"ReferencedDocument.docx");
        field->set_IsPathRelative(true);

        ASSERT_EQ(u" RD  ReferencedDocument.docx \\f", field->GetFieldCode());

        // Create the document that the RD field is referencing and insert a heading.
        // This heading will show up as an entry in the TOC field in our first document.
        auto referencedDoc = MakeObject<Document>();
        auto refDocBuilder = MakeObject<DocumentBuilder>(referencedDoc);
        refDocBuilder->get_CurrentParagraph()->get_ParagraphFormat()->set_StyleName(u"Heading 1");
        refDocBuilder->Writeln(u"TOC entry from referenced document");
        referencedDoc->Save(ArtifactsDir + u"ReferencedDocument.docx");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.RD.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.RD.docx");

        auto fieldToc = System::DynamicCast<FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(String(u"TOC entry from within this document\t\u0013 PAGEREF _Toc36149519 \\h \u00142\u0015\r") + u"TOC entry from referenced document\t1\r",
                  fieldToc->get_Result());

        auto fieldPageRef = System::DynamicCast<FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF _Toc36149519 \\h ", u"2", fieldPageRef);

        field = System::DynamicCast<FieldRD>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldRefDoc, u" RD  ReferencedDocument.docx \\f", String::Empty, field);
        ASSERT_EQ(u"ReferencedDocument.docx", field->get_FileName());
        ASSERT_TRUE(field->get_IsPathRelative());
    }

    void FieldSetRef()
    {
        //ExStart
        //ExFor:FieldRef
        //ExFor:FieldRef.BookmarkName
        //ExFor:FieldSet
        //ExFor:FieldSet.BookmarkName
        //ExFor:FieldSet.BookmarkText
        //ExSummary:Shows how to create bookmarked text with a SET field, and then display it in the document using a REF field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Name bookmarked text with a SET field.
        // This field refers to the "bookmark" not a bookmark structure that appears within the text, but a named variable.
        auto fieldSet = System::DynamicCast<FieldSet>(builder->InsertField(FieldType::FieldSet, false));
        fieldSet->set_BookmarkName(u"MyBookmark");
        fieldSet->set_BookmarkText(u"Hello world!");
        fieldSet->Update();

        ASSERT_EQ(u" SET  MyBookmark \"Hello world!\"", fieldSet->GetFieldCode());

        // Refer to the bookmark by name in a REF field and display its contents.
        auto fieldRef = System::DynamicCast<FieldRef>(builder->InsertField(FieldType::FieldRef, true));
        fieldRef->set_BookmarkName(u"MyBookmark");
        fieldRef->Update();

        ASSERT_EQ(u" REF  MyBookmark", fieldRef->GetFieldCode());
        ASSERT_EQ(u"Hello world!", fieldRef->get_Result());

        doc->Save(ArtifactsDir + u"Field.SET.REF.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.SET.REF.docx");

        ASSERT_EQ(u"Hello world!", doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Text());

        fieldSet = System::DynamicCast<FieldSet>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldSet, u" SET  MyBookmark \"Hello world!\"", u"Hello world!", fieldSet);
        ASSERT_EQ(u"MyBookmark", fieldSet->get_BookmarkName());
        ASSERT_EQ(u"Hello world!", fieldSet->get_BookmarkText());

        TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark", u"Hello world!", fieldRef);
        ASSERT_EQ(u"Hello world!", fieldRef->get_Result());
    }

    void FieldTemplate_()
    {
        //ExStart
        //ExFor:FieldTemplate
        //ExFor:FieldTemplate.IncludeFullPath
        //ExFor:FieldOptions.TemplateName
        //ExSummary:Shows how to use a TEMPLATE field to display the local file system location of a document's template.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // We can set a template name using by the fields. This property is used when the "doc.AttachedTemplate" is empty.
        // If this property is empty the default template file name "Normal.dotm" is used.
        doc->get_FieldOptions()->set_TemplateName(String::Empty);

        auto field = System::DynamicCast<FieldTemplate>(builder->InsertField(FieldType::FieldTemplate, false));
        ASSERT_EQ(u" TEMPLATE ", field->GetFieldCode());

        builder->Writeln();
        field = System::DynamicCast<FieldTemplate>(builder->InsertField(FieldType::FieldTemplate, false));
        field->set_IncludeFullPath(true);

        ASSERT_EQ(u" TEMPLATE  \\p", field->GetFieldCode());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.TEMPLATE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.TEMPLATE.docx");

        field = System::DynamicCast<FieldTemplate>(doc->get_Range()->get_Fields()->idx_get(0));
        ASSERT_EQ(u" TEMPLATE ", field->GetFieldCode());
        ASSERT_EQ(u"Normal.dotm", field->get_Result());

        field = System::DynamicCast<FieldTemplate>(doc->get_Range()->get_Fields()->idx_get(1));
        ASSERT_EQ(u" TEMPLATE  \\p", field->GetFieldCode());
        ASSERT_EQ(u"Normal.dotm", field->get_Result());
    }

    void FieldSymbol_()
    {
        //ExStart
        //ExFor:FieldSymbol
        //ExFor:FieldSymbol.CharacterCode
        //ExFor:FieldSymbol.DontAffectsLineSpacing
        //ExFor:FieldSymbol.FontName
        //ExFor:FieldSymbol.FontSize
        //ExFor:FieldSymbol.IsAnsi
        //ExFor:FieldSymbol.IsShiftJis
        //ExFor:FieldSymbol.IsUnicode
        //ExSummary:Shows how to use the SYMBOL field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are three ways to use a SYMBOL field to display a single character.
        // 1 -  Add a SYMBOL field which displays the © (Copyright) symbol, specified by an ANSI character code:
        auto field = System::DynamicCast<FieldSymbol>(builder->InsertField(FieldType::FieldSymbol, true));

        // The ANSI character code "U+00A9", or "169" in integer form, is reserved for the copyright symbol.
        field->set_CharacterCode(System::Convert::ToString(0x00a9));
        field->set_IsAnsi(true);

        ASSERT_EQ(u" SYMBOL  169 \\a", field->GetFieldCode());

        builder->Writeln(u" Line 1");

        // 2 -  Add a SYMBOL field which displays the ∞ (Infinity) symbol, and modify its appearance:
        field = System::DynamicCast<FieldSymbol>(builder->InsertField(FieldType::FieldSymbol, true));

        // In Unicode, the infinity symbol occupies the "221E" code.
        field->set_CharacterCode(System::Convert::ToString(0x221E));
        field->set_IsUnicode(true);

        // Change the font of our symbol after using the Windows Character Map
        // to ensure that the font can represent that symbol.
        field->set_FontName(u"Calibri");
        field->set_FontSize(u"24");

        // We can set this flag for tall symbols to make them not push down the rest of the text on their line.
        field->set_DontAffectsLineSpacing(true);

        ASSERT_EQ(u" SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", field->GetFieldCode());

        builder->Writeln(u"Line 2");

        // 3 -  Add a SYMBOL field which displays the あ character,
        // with a font that supports Shift-JIS (Windows-932) codepage:
        field = System::DynamicCast<FieldSymbol>(builder->InsertField(FieldType::FieldSymbol, true));
        field->set_FontName(u"MS Gothic");
        field->set_CharacterCode(System::Convert::ToString(0x82A0));
        field->set_IsShiftJis(true);

        ASSERT_EQ(u" SYMBOL  33440 \\f \"MS Gothic\" \\j", field->GetFieldCode());

        builder->Write(u"Line 3");

        doc->Save(ArtifactsDir + u"Field.SYMBOL.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.SYMBOL.docx");

        field = System::DynamicCast<FieldSymbol>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldSymbol, u" SYMBOL  169 \\a", String::Empty, field);
        ASSERT_EQ(System::Convert::ToString(0x00a9), field->get_CharacterCode());
        ASSERT_TRUE(field->get_IsAnsi());
        ASSERT_EQ(u"©", field->get_DisplayResult());

        field = System::DynamicCast<FieldSymbol>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldSymbol, u" SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", String::Empty, field);
        ASSERT_EQ(System::Convert::ToString(0x221E), field->get_CharacterCode());
        ASSERT_EQ(u"Calibri", field->get_FontName());
        ASSERT_EQ(u"24", field->get_FontSize());
        ASSERT_TRUE(field->get_IsUnicode());
        ASSERT_TRUE(field->get_DontAffectsLineSpacing());
        ASSERT_EQ(u"∞", field->get_DisplayResult());

        field = System::DynamicCast<FieldSymbol>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldSymbol, u" SYMBOL  33440 \\f \"MS Gothic\" \\j", String::Empty, field);
        ASSERT_EQ(System::Convert::ToString(0x82A0), field->get_CharacterCode());
        ASSERT_EQ(u"MS Gothic", field->get_FontName());
        ASSERT_TRUE(field->get_IsShiftJis());
    }

    void FieldTitle_()
    {
        //ExStart
        //ExFor:FieldTitle
        //ExFor:FieldTitle.Text
        //ExSummary:Shows how to use the TITLE field.
        auto doc = MakeObject<Document>();

        // Set a value for the "Title" built-in document property.
        doc->get_BuiltInDocumentProperties()->set_Title(u"My Title");

        // We can use the TITLE field to display the value of this property in the document.
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto field = System::DynamicCast<FieldTitle>(builder->InsertField(FieldType::FieldTitle, false));
        field->Update();

        ASSERT_EQ(u" TITLE ", field->GetFieldCode());
        ASSERT_EQ(u"My Title", field->get_Result());

        // Setting a value for the field's Text property,
        // and then updating the field will also overwrite the corresponding built-in property with the new value.
        builder->Writeln();
        field = System::DynamicCast<FieldTitle>(builder->InsertField(FieldType::FieldTitle, false));
        field->set_Text(u"My New Title");
        field->Update();

        ASSERT_EQ(u" TITLE  \"My New Title\"", field->GetFieldCode());
        ASSERT_EQ(u"My New Title", field->get_Result());
        ASSERT_EQ(u"My New Title", doc->get_BuiltInDocumentProperties()->get_Title());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.TITLE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.TITLE.docx");

        ASSERT_EQ(u"My New Title", doc->get_BuiltInDocumentProperties()->get_Title());

        field = System::DynamicCast<FieldTitle>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldTitle, u" TITLE ", u"My New Title", field);

        field = System::DynamicCast<FieldTitle>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldTitle, u" TITLE  \"My New Title\"", u"My New Title", field);
        ASSERT_EQ(u"My New Title", field->get_Text());
    }

    //ExStart
    //ExFor:FieldToa
    //ExFor:FieldToa.BookmarkName
    //ExFor:FieldToa.EntryCategory
    //ExFor:FieldToa.EntrySeparator
    //ExFor:FieldToa.PageNumberListSeparator
    //ExFor:FieldToa.PageRangeSeparator
    //ExFor:FieldToa.RemoveEntryFormatting
    //ExFor:FieldToa.SequenceName
    //ExFor:FieldToa.SequenceSeparator
    //ExFor:FieldToa.UseHeading
    //ExFor:FieldToa.UsePassim
    //ExFor:FieldTA
    //ExFor:FieldTA.EntryCategory
    //ExFor:FieldTA.IsBold
    //ExFor:FieldTA.IsItalic
    //ExFor:FieldTA.LongCitation
    //ExFor:FieldTA.PageRangeBookmarkName
    //ExFor:FieldTA.ShortCitation
    //ExSummary:Shows how to build and customize a table of authorities using TOA and TA fields.
    void FieldTOA_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a TOA field, which will create an entry for each TA field in the document,
        // displaying long citations and page numbers for each entry.
        auto fieldToa = System::DynamicCast<FieldToa>(builder->InsertField(FieldType::FieldTOA, false));

        // Set the entry category for our table. This TOA will now only include TA fields
        // that have a matching value in their EntryCategory property.
        fieldToa->set_EntryCategory(u"1");

        // Moreover, the Table of Authorities category at index 1 is "Cases",
        // which will show up as our table's title if we set this variable to true.
        fieldToa->set_UseHeading(true);

        // We can further filter TA fields by naming a bookmark that they will need to be within the TOA bounds.
        fieldToa->set_BookmarkName(u"MyBookmark");

        // By default, a dotted line page-wide tab appears between the TA field's citation
        // and its page number. We can replace it with any text we put on this property.
        // Inserting a tab character will preserve the original tab.
        fieldToa->set_EntrySeparator(u" \t p.");

        // If we have multiple TA entries that share the same long citation,
        // all their respective page numbers will show up on one row.
        // We can use this property to specify a string that will separate their page numbers.
        fieldToa->set_PageNumberListSeparator(u" & p. ");

        // We can set this to true to get our table to display the word "passim"
        // if there are five or more page numbers in one row.
        fieldToa->set_UsePassim(true);

        // One TA field can refer to a range of pages.
        // We can specify a string here to appear between the start and end page numbers for such ranges.
        fieldToa->set_PageRangeSeparator(u" to ");

        // The format from the TA fields will carry over into our table.
        // We can disable this by setting the RemoveEntryFormatting flag.
        fieldToa->set_RemoveEntryFormatting(true);
        builder->get_Font()->set_Color(System::Drawing::Color::get_Green());
        builder->get_Font()->set_Name(u"Arial Black");

        ASSERT_EQ(u" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f", fieldToa->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);

        // This TA field will not appear as an entry in the TOA since it is outside
        // the bookmark's bounds that the TOA's BookmarkName property specifies.
        SharedPtr<FieldTA> fieldTA = InsertToaEntry(builder, u"1", u"Source 1");

        ASSERT_EQ(u" TA  \\c 1 \\l \"Source 1\"", fieldTA->GetFieldCode());

        // This TA field is inside the bookmark,
        // but the entry category does not match that of the table, so the TA field will not include it.
        builder->StartBookmark(u"MyBookmark");
        fieldTA = InsertToaEntry(builder, u"2", u"Source 2");

        // This entry will appear in the table.
        fieldTA = InsertToaEntry(builder, u"1", u"Source 3");

        // A TOA table does not display short citations,
        // but we can use them as a shorthand to refer to bulky source names that multiple TA fields reference.
        fieldTA->set_ShortCitation(u"S.3");

        ASSERT_EQ(u" TA  \\c 1 \\l \"Source 3\" \\s S.3", fieldTA->GetFieldCode());

        // We can format the page number to make it bold/italic using the following properties.
        // We will still see these effects if we set our table to ignore formatting.
        fieldTA = InsertToaEntry(builder, u"1", u"Source 2");
        fieldTA->set_IsBold(true);
        fieldTA->set_IsItalic(true);

        ASSERT_EQ(u" TA  \\c 1 \\l \"Source 2\" \\b \\i", fieldTA->GetFieldCode());

        // We can configure TA fields to get their TOA entries to refer to a range of pages that a bookmark spans across.
        // Note that this entry refers to the same source as the one above to share one row in our table.
        // This row will have the page number of the entry above and the page range of this entry,
        // with the table's page list and page number range separators between page numbers.
        fieldTA = InsertToaEntry(builder, u"1", u"Source 3");
        fieldTA->set_PageRangeBookmarkName(u"MyMultiPageBookmark");

        builder->StartBookmark(u"MyMultiPageBookmark");
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::PageBreak);
        builder->EndBookmark(u"MyMultiPageBookmark");

        ASSERT_EQ(u" TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", fieldTA->GetFieldCode());

        // If we have enabled the "Passim" feature of our table, having 5 or more TA entries with the same source will invoke it.
        for (int i = 0; i < 5; i++)
        {
            InsertToaEntry(builder, u"1", u"Source 4");
        }

        builder->EndBookmark(u"MyBookmark");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.TOA.TA.docx");
        TestFieldTOA(MakeObject<Document>(ArtifactsDir + u"Field.TOA.TA.docx"));
        //ExSkip
    }

    static SharedPtr<FieldTA> InsertToaEntry(SharedPtr<DocumentBuilder> builder, String entryCategory, String longCitation)
    {
        auto field = System::DynamicCast<FieldTA>(builder->InsertField(FieldType::FieldTOAEntry, false));
        field->set_EntryCategory(entryCategory);
        field->set_LongCitation(longCitation);

        builder->InsertBreak(BreakType::PageBreak);

        return field;
    }
    //ExEnd

    void TestFieldTOA(SharedPtr<Document> doc)
    {
        auto fieldTOA = System::DynamicCast<FieldToa>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(u"1", fieldTOA->get_EntryCategory());
        ASSERT_TRUE(fieldTOA->get_UseHeading());
        ASSERT_EQ(u"MyBookmark", fieldTOA->get_BookmarkName());
        ASSERT_EQ(u" \t p.", fieldTOA->get_EntrySeparator());
        ASSERT_EQ(u" & p. ", fieldTOA->get_PageNumberListSeparator());
        ASSERT_TRUE(fieldTOA->get_UsePassim());
        ASSERT_EQ(u" to ", fieldTOA->get_PageRangeSeparator());
        ASSERT_TRUE(fieldTOA->get_RemoveEntryFormatting());
        ASSERT_EQ(u" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f", fieldTOA->GetFieldCode());
        ASSERT_EQ(String(u"Cases\r") + u"Source 2 \t p.5\r" + u"Source 3 \t p.4 & p. 7 to 10\r" + u"Source 4 \t p.passim\r", fieldTOA->get_Result());

        auto fieldTA = System::DynamicCast<FieldTA>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldTOAEntry, u" TA  \\c 1 \\l \"Source 1\"", String::Empty, fieldTA);
        ASSERT_EQ(u"1", fieldTA->get_EntryCategory());
        ASSERT_EQ(u"Source 1", fieldTA->get_LongCitation());

        fieldTA = System::DynamicCast<FieldTA>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldTOAEntry, u" TA  \\c 2 \\l \"Source 2\"", String::Empty, fieldTA);
        ASSERT_EQ(u"2", fieldTA->get_EntryCategory());
        ASSERT_EQ(u"Source 2", fieldTA->get_LongCitation());

        fieldTA = System::DynamicCast<FieldTA>(doc->get_Range()->get_Fields()->idx_get(3));

        TestUtil::VerifyField(FieldType::FieldTOAEntry, u" TA  \\c 1 \\l \"Source 3\" \\s S.3", String::Empty, fieldTA);
        ASSERT_EQ(u"1", fieldTA->get_EntryCategory());
        ASSERT_EQ(u"Source 3", fieldTA->get_LongCitation());
        ASSERT_EQ(u"S.3", fieldTA->get_ShortCitation());

        fieldTA = System::DynamicCast<FieldTA>(doc->get_Range()->get_Fields()->idx_get(4));

        TestUtil::VerifyField(FieldType::FieldTOAEntry, u" TA  \\c 1 \\l \"Source 2\" \\b \\i", String::Empty, fieldTA);
        ASSERT_EQ(u"1", fieldTA->get_EntryCategory());
        ASSERT_EQ(u"Source 2", fieldTA->get_LongCitation());
        ASSERT_TRUE(fieldTA->get_IsBold());
        ASSERT_TRUE(fieldTA->get_IsItalic());

        fieldTA = System::DynamicCast<FieldTA>(doc->get_Range()->get_Fields()->idx_get(5));

        TestUtil::VerifyField(FieldType::FieldTOAEntry, u" TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", String::Empty, fieldTA);
        ASSERT_EQ(u"1", fieldTA->get_EntryCategory());
        ASSERT_EQ(u"Source 3", fieldTA->get_LongCitation());
        ASSERT_EQ(u"MyMultiPageBookmark", fieldTA->get_PageRangeBookmarkName());

        for (int i = 6; i < 11; i++)
        {
            fieldTA = System::DynamicCast<FieldTA>(doc->get_Range()->get_Fields()->idx_get(i));

            TestUtil::VerifyField(FieldType::FieldTOAEntry, u" TA  \\c 1 \\l \"Source 4\"", String::Empty, fieldTA);
            ASSERT_EQ(u"1", fieldTA->get_EntryCategory());
            ASSERT_EQ(u"Source 4", fieldTA->get_LongCitation());
        }
    }

    void FieldAddIn_()
    {
        //ExStart
        //ExFor:FieldAddIn
        //ExSummary:Shows how to process an ADDIN field.
        auto doc = MakeObject<Document>(MyDir + u"Field sample - ADDIN.docx");

        // Aspose.Words does not support inserting ADDIN fields, but we can still load and read them.
        auto field = System::DynamicCast<FieldAddIn>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(u" ADDIN \"My value\" ", field->GetFieldCode());
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);

        TestUtil::VerifyField(FieldType::FieldAddin, u" ADDIN \"My value\" ", String::Empty, doc->get_Range()->get_Fields()->idx_get(0));
    }

    void FieldEditTime_()
    {
        //ExStart
        //ExFor:FieldEditTime
        //ExSummary:Shows how to use the EDITTIME field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // The EDITTIME field will show, in minutes,
        // the time spent with the document open in a Microsoft Word window.
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Write(u"You've been editing this document for ");
        auto field = System::DynamicCast<FieldEditTime>(builder->InsertField(FieldType::FieldEditTime, true));
        builder->Writeln(u" minutes.");

        // This built in document property tracks the minutes. Microsoft Word uses this property
        // to track the time spent with the document open. We can also edit it ourselves.
        doc->get_BuiltInDocumentProperties()->set_TotalEditingTime(10);
        field->Update();

        ASSERT_EQ(u" EDITTIME ", field->GetFieldCode());
        ASSERT_EQ(u"10", field->get_Result());

        // The field does not update itself in real-time, and will also have to be
        // manually updated in Microsoft Word anytime we need an accurate value.
        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.EDITTIME.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.EDITTIME.docx");

        ASSERT_EQ(10, doc->get_BuiltInDocumentProperties()->get_TotalEditingTime());

        TestUtil::VerifyField(FieldType::FieldEditTime, u" EDITTIME ", u"10", doc->get_Range()->get_Fields()->idx_get(0));
    }

    //ExStart
    //ExFor:FieldEQ
    //ExSummary:Shows how to use the EQ field to display a variety of mathematical equations.
    void FieldEQ_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // An EQ field displays a mathematical equation consisting of one or many elements.
        // Each element takes the following form: [switch][options][arguments].
        // There may be one switch, and several possible options.
        // The arguments are a set of coma-separated values enclosed by round braces.

        // Here we use a document builder to insert an EQ field, with an "\f" switch, which corresponds to "Fraction".
        // We will pass values 1 and 4 as arguments, and we will not use any options.
        // This field will display a fraction with 1 as the numerator and 4 as the denominator.
        SharedPtr<FieldEQ> field = InsertFieldEQ(builder, u"\\f(1,4)");

        ASSERT_EQ(u" EQ \\f(1,4)", field->GetFieldCode());

        // One EQ field may contain multiple elements placed sequentially.
        // We can also nest elements inside one another by placing the inner elements
        // inside the argument brackets of outer elements.
        // We can find the full list of switches, along with their uses here:
        // https://blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/

        // Below are applications of nine different EQ field switches that we can use to create different kinds of objects.
        // 1 -  Array switch "\a", aligned left, 2 columns, 3 points of horizontal and vertical spacing:
        InsertFieldEQ(builder, u"\\a \\al \\co2 \\vs3 \\hs3(4x,- 4y,-4x,+ y)");

        // 2 -  Bracket switch "\b", bracket character "[", to enclose the contents in a set of square braces:
        // Note that we are nesting an array inside the brackets, which will altogether look like a matrix in the output.
        InsertFieldEQ(builder, u"\\b \\bc\\[ (\\a \\al \\co3 \\vs3 \\hs3(1,0,0,0,1,0,0,0,1))");

        // 3 -  Displacement switch "\d", displacing text "B" 30 spaces to the right of "A", displaying the gap as an underline:
        InsertFieldEQ(builder, u"A \\d \\fo30 \\li() B");

        // 4 -  Formula consisting of multiple fractions:
        InsertFieldEQ(builder, u"\\f(d,dx)(u + v) = \\f(du,dx) + \\f(dv,dx)");

        // 5 -  Integral switch "\i", with a summation symbol:
        InsertFieldEQ(builder, u"\\i \\su(n=1,5,n)");

        // 6 -  List switch "\l":
        InsertFieldEQ(builder, u"\\l(1,1,2,3,n,8,13)");

        // 7 -  Radical switch "\r", displaying a cubed root of x:
        InsertFieldEQ(builder, u"\\r (3,x)");

        // 8 -  Subscript/superscript switch "/s", first as a superscript and then as a subscript:
        InsertFieldEQ(builder, u"\\s \\up8(Superscript) Text \\s \\do8(Subscript)");

        // 9 -  Box switch "\x", with lines at the top, bottom, left and right of the input:
        InsertFieldEQ(builder, u"\\x \\to \\bo \\le \\ri(5)");

        // Some more complex combinations.
        InsertFieldEQ(builder, u"\\a \\ac \\vs1 \\co1(lim,n→∞) \\b (\\f(n,n2 + 12) + \\f(n,n2 + 22) + ... + \\f(n,n2 + n2))");
        InsertFieldEQ(builder, u"\\i (,,  \\b(\\f(x,x2 + 3x + 2))) \\s \\up10(2)");
        InsertFieldEQ(builder, u"\\i \\in( tan x, \\s \\up2(sec x), \\b(\\r(3) )\\s \\up4(t) \\s \\up7(2)  dt)");

        doc->Save(ArtifactsDir + u"Field.EQ.docx");
        TestFieldEQ(MakeObject<Document>(ArtifactsDir + u"Field.EQ.docx"));
        //ExSkip
    }

    /// <summary>
    /// Use a document builder to insert an EQ field, set its arguments and start a new paragraph.
    /// </summary>
    static SharedPtr<FieldEQ> InsertFieldEQ(SharedPtr<DocumentBuilder> builder, String args)
    {
        auto field = System::DynamicCast<FieldEQ>(builder->InsertField(FieldType::FieldEquation, true));
        builder->MoveTo(field->get_Separator());
        builder->Write(args);
        builder->MoveTo(field->get_Start()->get_ParentNode());

        builder->InsertParagraph();
        return field;
    }
    //ExEnd

    void TestFieldEQ(SharedPtr<Document> doc)
    {
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\f(1,4)", String::Empty, doc->get_Range()->get_Fields()->idx_get(0));
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\a \\al \\co2 \\vs3 \\hs3(4x,- 4y,-4x,+ y)", String::Empty,
                              doc->get_Range()->get_Fields()->idx_get(1));
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\b \\bc\\[ (\\a \\al \\co3 \\vs3 \\hs3(1,0,0,0,1,0,0,0,1))", String::Empty,
                              doc->get_Range()->get_Fields()->idx_get(2));
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ A \\d \\fo30 \\li() B", String::Empty, doc->get_Range()->get_Fields()->idx_get(3));
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\f(d,dx)(u + v) = \\f(du,dx) + \\f(dv,dx)", String::Empty,
                              doc->get_Range()->get_Fields()->idx_get(4));
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\i \\su(n=1,5,n)", String::Empty, doc->get_Range()->get_Fields()->idx_get(5));
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\l(1,1,2,3,n,8,13)", String::Empty, doc->get_Range()->get_Fields()->idx_get(6));
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\r (3,x)", String::Empty, doc->get_Range()->get_Fields()->idx_get(7));
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\s \\up8(Superscript) Text \\s \\do8(Subscript)", String::Empty,
                              doc->get_Range()->get_Fields()->idx_get(8));
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\x \\to \\bo \\le \\ri(5)", String::Empty, doc->get_Range()->get_Fields()->idx_get(9));
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\a \\ac \\vs1 \\co1(lim,n→∞) \\b (\\f(n,n2 + 12) + \\f(n,n2 + 22) + ... + \\f(n,n2 + n2))",
                              String::Empty, doc->get_Range()->get_Fields()->idx_get(10));
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\i (,,  \\b(\\f(x,x2 + 3x + 2))) \\s \\up10(2)", String::Empty,
                              doc->get_Range()->get_Fields()->idx_get(11));
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\i \\in( tan x, \\s \\up2(sec x), \\b(\\r(3) )\\s \\up4(t) \\s \\up7(2)  dt)", String::Empty,
                              doc->get_Range()->get_Fields()->idx_get(12));
        TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, u"https://blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/");
    }

    void FieldForms()
    {
        //ExStart
        //ExFor:FieldFormCheckBox
        //ExFor:FieldFormDropDown
        //ExFor:FieldFormText
        //ExSummary:Shows how to process FORMCHECKBOX, FORMDROPDOWN and FORMTEXT fields.
        // These fields are legacy equivalents of the FormField. We can read, but not create these fields using Aspose.Words.
        // In Microsoft Word, we can insert these fields via the Legacy Tools menu in the Developer tab.
        auto doc = MakeObject<Document>(MyDir + u"Form fields.docx");

        auto fieldFormCheckBox = System::DynamicCast<FieldFormCheckBox>(doc->get_Range()->get_Fields()->idx_get(1));
        ASSERT_EQ(u" FORMCHECKBOX \u0001", fieldFormCheckBox->GetFieldCode());

        auto fieldFormDropDown = System::DynamicCast<FieldFormDropDown>(doc->get_Range()->get_Fields()->idx_get(2));
        ASSERT_EQ(u" FORMDROPDOWN \u0001", fieldFormDropDown->GetFieldCode());

        auto fieldFormText = System::DynamicCast<FieldFormText>(doc->get_Range()->get_Fields()->idx_get(0));
        ASSERT_EQ(u" FORMTEXT \u0001", fieldFormText->GetFieldCode());
        //ExEnd
    }

    void FieldFormula_()
    {
        //ExStart
        //ExFor:FieldFormula
        //ExSummary:Shows how to use the formula field to display the result of an equation.
        auto doc = MakeObject<Document>();

        // Use a field builder to construct a mathematical equation,
        // then create a formula field to display the equation's result in the document.
        auto fieldBuilder = MakeObject<FieldBuilder>(FieldType::FieldFormula);
        fieldBuilder->AddArgument(2);
        fieldBuilder->AddArgument(u"*");
        fieldBuilder->AddArgument(5);

        auto field = System::DynamicCast<FieldFormula>(fieldBuilder->BuildAndInsert(doc->get_FirstSection()->get_Body()->get_FirstParagraph()));
        field->Update();

        ASSERT_EQ(u" = 2 * 5 ", field->GetFieldCode());
        ASSERT_EQ(u"10", field->get_Result());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.FORMULA.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.FORMULA.docx");

        TestUtil::VerifyField(FieldType::FieldFormula, u" = 2 * 5 ", u"10", doc->get_Range()->get_Fields()->idx_get(0));
    }

    void FieldLastSavedBy_()
    {
        //ExStart
        //ExFor:FieldLastSavedBy
        //ExSummary:Shows how to use the LASTSAVEDBY field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // If we create a document in Microsoft Word, it will have the user's name in the "Last saved by" built-in property.
        // If we make a document programmatically, this property will be null, and we will need to assign a value.
        doc->get_BuiltInDocumentProperties()->set_LastSavedBy(u"John Doe");

        // We can use the LASTSAVEDBY field to display the value of this property in the document.
        auto field = System::DynamicCast<FieldLastSavedBy>(builder->InsertField(FieldType::FieldLastSavedBy, true));

        ASSERT_EQ(u" LASTSAVEDBY ", field->GetFieldCode());
        ASSERT_EQ(u"John Doe", field->get_Result());

        doc->Save(ArtifactsDir + u"Field.LASTSAVEDBY.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.LASTSAVEDBY.docx");

        ASSERT_EQ(u"John Doe", doc->get_BuiltInDocumentProperties()->get_LastSavedBy());
        TestUtil::VerifyField(FieldType::FieldLastSavedBy, u" LASTSAVEDBY ", u"John Doe", doc->get_Range()->get_Fields()->idx_get(0));
    }

    void FieldOcx_()
    {
        //ExStart
        //ExFor:FieldOcx
        //ExSummary:Shows how to insert an OCX field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto field = System::DynamicCast<FieldOcx>(builder->InsertField(FieldType::FieldOcx, true));

        ASSERT_EQ(u" OCX ", field->GetFieldCode());
        //ExEnd

        TestUtil::VerifyField(FieldType::FieldOcx, u" OCX ", String::Empty, field);
    }

    //ExStart
    //ExFor:Field.Remove
    //ExFor:FieldPrivate
    //ExSummary:Shows how to process PRIVATE fields.
    void FieldPrivate_()
    {
        // Open a Corel WordPerfect document which we have converted to .docx format.
        auto doc = MakeObject<Document>(MyDir + u"Field sample - PRIVATE.docx");

        // WordPerfect 5.x/6.x documents like the one we have loaded may contain PRIVATE fields.
        // Microsoft Word preserves PRIVATE fields during load/save operations,
        // but provides no functionality for them.
        auto field = System::DynamicCast<FieldPrivate>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(u" PRIVATE \"My value\" ", field->GetFieldCode());
        ASSERT_EQ(FieldType::FieldPrivate, field->get_Type());

        // We can also insert PRIVATE fields using a document builder.
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertField(FieldType::FieldPrivate, true);

        // These fields are not a viable way of protecting sensitive information.
        // Unless backward compatibility with older versions of WordPerfect is essential,
        // we can safely remove these fields. We can do this using a DocumentVisiitor implementation.
        ASSERT_EQ(2, doc->get_Range()->get_Fields()->get_Count());

        auto remover = MakeObject<ExField::FieldPrivateRemover>();
        doc->Accept(remover);

        ASSERT_EQ(2, remover->GetFieldsRemovedCount());
        ASSERT_EQ(0, doc->get_Range()->get_Fields()->get_Count());
    }

    /// <summary>
    /// Removes all encountered PRIVATE fields.
    /// </summary>
    class FieldPrivateRemover : public DocumentVisitor
    {
    public:
        FieldPrivateRemover() : mFieldsRemovedCount(0)
        {
            mFieldsRemovedCount = 0;
        }

        int GetFieldsRemovedCount()
        {
            return mFieldsRemovedCount;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// If the node belongs to a PRIVATE field, the entire field is removed.
        /// </summary>
        VisitorAction VisitFieldEnd(SharedPtr<FieldEnd> fieldEnd) override
        {
            if (fieldEnd->get_FieldType() == FieldType::FieldPrivate)
            {
                fieldEnd->GetField()->Remove();
                mFieldsRemovedCount++;
            }

            return VisitorAction::Continue;
        }

    private:
        int mFieldsRemovedCount;
    };
    //ExEnd

    void FieldSection_()
    {
        //ExStart
        //ExFor:FieldSection
        //ExFor:FieldSectionPages
        //ExSummary:Shows how to use SECTION and SECTIONPAGES fields to number pages by sections.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);

        // A SECTION field displays the number of the section it is in.
        builder->Write(u"Section ");
        auto fieldSection = System::DynamicCast<FieldSection>(builder->InsertField(FieldType::FieldSection, true));

        ASSERT_EQ(u" SECTION ", fieldSection->GetFieldCode());

        // A PAGE field displays the number of the page it is in.
        builder->Write(u"\nPage ");
        auto fieldPage = System::DynamicCast<FieldPage>(builder->InsertField(FieldType::FieldPage, true));

        ASSERT_EQ(u" PAGE ", fieldPage->GetFieldCode());

        // A SECTIONPAGES field displays the number of pages that the section it is in spans across.
        builder->Write(u" of ");
        auto fieldSectionPages = System::DynamicCast<FieldSectionPages>(builder->InsertField(FieldType::FieldSectionPages, true));

        ASSERT_EQ(u" SECTIONPAGES ", fieldSectionPages->GetFieldCode());

        // Move out of the header back into the main document and insert two pages.
        // All these pages will be in the first section. Our fields, which appear once every header,
        // will number the current/total pages of this section.
        builder->MoveToDocumentEnd();
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::PageBreak);

        // We can insert a new section with the document builder like this.
        // This will affect the values displayed in the SECTION and SECTIONPAGES fields in all upcoming headers.
        builder->InsertBreak(BreakType::SectionBreakNewPage);

        // The PAGE field will keep counting pages across the whole document.
        // We can manually reset its count at each section to keep track of pages section-by-section.
        builder->get_CurrentSection()->get_PageSetup()->set_RestartPageNumbering(true);
        builder->InsertBreak(BreakType::PageBreak);

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.SECTION.SECTIONPAGES.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.SECTION.SECTIONPAGES.docx");

        TestUtil::VerifyField(FieldType::FieldSection, u" SECTION ", u"2", doc->get_Range()->get_Fields()->idx_get(0));
        TestUtil::VerifyField(FieldType::FieldPage, u" PAGE ", u"2", doc->get_Range()->get_Fields()->idx_get(1));
        TestUtil::VerifyField(FieldType::FieldSectionPages, u" SECTIONPAGES ", u"2", doc->get_Range()->get_Fields()->idx_get(2));
    }

    //ExStart
    //ExFor:FieldTime
    //ExSummary:Shows how to display the current time using the TIME field.
    void FieldTime_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // By default, time is displayed in the "h:mm am/pm" format.
        SharedPtr<FieldTime> field = InsertFieldTime(builder, u"");

        ASSERT_EQ(u" TIME ", field->GetFieldCode());

        // We can use the \@ flag to change the format of our displayed time.
        field = InsertFieldTime(builder, u"\\@ HHmm");

        ASSERT_EQ(u" TIME \\@ HHmm", field->GetFieldCode());

        // We can adjust the format to get TIME field to also display the date, according to the Gregorian calendar.
        field = InsertFieldTime(builder, u"\\@ \"M/d/yyyy h mm:ss am/pm\"");

        ASSERT_EQ(u" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"", field->GetFieldCode());

        doc->Save(ArtifactsDir + u"Field.TIME.docx");
        TestFieldTime(MakeObject<Document>(ArtifactsDir + u"Field.TIME.docx"));
        //ExSkip
    }

    /// <summary>
    /// Use a document builder to insert a TIME field, insert a new paragraph and return the field.
    /// </summary>
    static SharedPtr<FieldTime> InsertFieldTime(SharedPtr<DocumentBuilder> builder, String format)
    {
        auto field = System::DynamicCast<FieldTime>(builder->InsertField(FieldType::FieldTime, true));
        builder->MoveTo(field->get_Separator());
        builder->Write(format);
        builder->MoveTo(field->get_Start()->get_ParentNode());

        builder->InsertParagraph();
        return field;
    }
    //ExEnd

    void TestFieldTime(SharedPtr<Document> doc)
    {
        System::DateTime docLoadingTime = System::DateTime::get_Now();
        doc = DocumentHelper::SaveOpen(doc);

        auto field = System::DynamicCast<FieldTime>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(u" TIME ", field->GetFieldCode());
        ASSERT_EQ(FieldType::FieldTime, field->get_Type());
        ASSERT_EQ(System::DateTime::Parse(field->get_Result()),
                  System::DateTime::get_Today().AddHours(docLoadingTime.get_Hour()).AddMinutes(docLoadingTime.get_Minute()));

        field = System::DynamicCast<FieldTime>(doc->get_Range()->get_Fields()->idx_get(1));

        ASSERT_EQ(u" TIME \\@ HHmm", field->GetFieldCode());
        ASSERT_EQ(FieldType::FieldTime, field->get_Type());
        ASSERT_EQ(System::DateTime::Parse(field->get_Result()),
                  System::DateTime::get_Today().AddHours(docLoadingTime.get_Hour()).AddMinutes(docLoadingTime.get_Minute()));

        field = System::DynamicCast<FieldTime>(doc->get_Range()->get_Fields()->idx_get(2));

        ASSERT_EQ(u" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"", field->GetFieldCode());
        ASSERT_EQ(FieldType::FieldTime, field->get_Type());
        ASSERT_EQ(System::DateTime::Parse(field->get_Result()),
                  System::DateTime::get_Today().AddHours(docLoadingTime.get_Hour()).AddMinutes(docLoadingTime.get_Minute()));
    }

    void BidiOutline()
    {
        //ExStart
        //ExFor:FieldBidiOutline
        //ExFor:FieldShape
        //ExFor:FieldShape.Text
        //ExFor:ParagraphFormat.Bidi
        //ExSummary:Shows how to create right-to-left language-compatible lists with BIDIOUTLINE fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // The BIDIOUTLINE field numbers paragraphs like the AUTONUM/LISTNUM fields,
        // but is only visible when a right-to-left editing language is enabled, such as Hebrew or Arabic.
        // The following field will display ".1", the RTL equivalent of list number "1.".
        auto field = System::DynamicCast<FieldBidiOutline>(builder->InsertField(FieldType::FieldBidiOutline, true));
        builder->Writeln(u"שלום");

        ASSERT_EQ(u" BIDIOUTLINE ", field->GetFieldCode());

        // Add two more BIDIOUTLINE fields, which will display ".2" and ".3".
        builder->InsertField(FieldType::FieldBidiOutline, true);
        builder->Writeln(u"שלום");
        builder->InsertField(FieldType::FieldBidiOutline, true);
        builder->Writeln(u"שלום");

        // Set the horizontal text alignment for every paragraph in the document to RTL.
        for (const auto& para : System::IterateOver<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)))
        {
            para->get_ParagraphFormat()->set_Bidi(true);
        }

        // If we enable a right-to-left editing language in Microsoft Word, our fields will display numbers.
        // Otherwise, they will display "###".
        doc->Save(ArtifactsDir + u"Field.BIDIOUTLINE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.BIDIOUTLINE.docx");

        for (const auto& fieldBidiOutline : System::IterateOver(doc->get_Range()->get_Fields()))
        {
            TestUtil::VerifyField(FieldType::FieldBidiOutline, u" BIDIOUTLINE ", String::Empty, fieldBidiOutline);
        }
    }

    void Legacy()
    {
        //ExStart
        //ExFor:FieldEmbed
        //ExFor:FieldShape
        //ExFor:FieldShape.Text
        //ExSummary:Shows how some older Microsoft Word fields such as SHAPE and EMBED are handled during loading.
        // Open a document that was created in Microsoft Word 2003.
        auto doc = MakeObject<Document>(MyDir + u"Legacy fields.doc");

        // If we open the Word document and press Alt+F9, we will see a SHAPE and an EMBED field.
        // A SHAPE field is the anchor/canvas for an AutoShape object with the "In line with text" wrapping style enabled.
        // An EMBED field has the same function, but for an embedded object,
        // such as a spreadsheet from an external Excel document.
        // However, these fields will not appear in the document's Fields collection.
        ASSERT_EQ(0, doc->get_Range()->get_Fields()->get_Count());

        // These fields are supported only by old versions of Microsoft Word.
        // The document loading process will convert these fields into Shape objects,
        // which we can access in the document's node collection.
        SharedPtr<NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);
        ASSERT_EQ(3, shapes->get_Count());

        // The first Shape node corresponds to the SHAPE field in the input document,
        // which is the inline canvas for the AutoShape.
        auto shape = System::DynamicCast<Shape>(shapes->idx_get(0));
        ASSERT_EQ(ShapeType::Image, shape->get_ShapeType());

        // The second Shape node is the AutoShape itself.
        shape = System::DynamicCast<Shape>(shapes->idx_get(1));
        ASSERT_EQ(ShapeType::Can, shape->get_ShapeType());

        // The third Shape is what was the EMBED field that contained the external spreadsheet.
        shape = System::DynamicCast<Shape>(shapes->idx_get(2));
        ASSERT_EQ(ShapeType::OleObject, shape->get_ShapeType());
        //ExEnd
    }

    void SetFieldIndexFormat()
    {
        //ExStart
        //ExFor:FieldOptions.FieldIndexFormat
        //ExSummary:Shows how to formatting FieldIndex fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u"A");
        builder->InsertBreak(BreakType::LineBreak);
        builder->InsertField(u"XE \"A\"");
        builder->Write(u"B");

        builder->InsertField(u" INDEX \\e \" · \" \\h \"A\" \\c \"2\" \\z \"1033\"", nullptr);

        doc->get_FieldOptions()->set_FieldIndexFormat(FieldIndexFormat::Fancy);
        doc->UpdateFields();

        doc->Save(ArtifactsDir + u"Field.SetFieldIndexFormat.docx");
        //ExEnd
    }

    //ExStart
    //ExFor:ComparisonEvaluationResult.#ctor(bool)
    //ExFor:ComparisonEvaluationResult.#ctor(string)
    //ExFor:ComparisonEvaluationResult
    //ExFor:ComparisonExpression
    //ExFor:ComparisonExpression.LeftExpression
    //ExFor:ComparisonExpression.ComparisonOperator
    //ExFor:ComparisonExpression.RightExpression
    //ExFor:FieldOptions.ComparisonExpressionEvaluator
    //ExSummary:Shows how to implement custom evaluation for the IF and COMPARE fields.
    void ConditionEvaluationExtensionPoint(String fieldCode, int8_t comparisonResult, String comparisonError, String expectedResult)
    {
        const String left = u"\"left expression\"";
        const String operator_ = u"<>";
        const String right = u"\"right expression\"";

        auto builder = MakeObject<DocumentBuilder>();

        // Field codes that we use in this example:
        // 1.   " IF {0} {1} {2} \"true argument\" \"false argument\" ".
        // 2.   " COMPARE {0} {1} {2} ".
        SharedPtr<Field> field = builder->InsertField(String::Format(fieldCode, left, operator_, right), nullptr);

        // If the "comparisonResult" is undefined, we create "ComparisonEvaluationResult" with string, instead of bool.
        SharedPtr<ComparisonEvaluationResult> result = comparisonResult != -1       ? MakeObject<ComparisonEvaluationResult>(comparisonResult == 1)
                                                       : comparisonError != nullptr ? MakeObject<ComparisonEvaluationResult>(comparisonError)
                                                                                    : nullptr;

        auto evaluator = MakeObject<ExField::ComparisonExpressionEvaluator>(result);
        builder->get_Document()->get_FieldOptions()->set_ComparisonExpressionEvaluator(evaluator);

        builder->get_Document()->UpdateFields();

        ASSERT_EQ(expectedResult, field->get_Result());
        evaluator->AssertInvocationsCount(1)->AssertInvocationArguments(0, left, operator_, right);
    }

    /// <summary>
    /// Comparison expressions evaluation for the FieldIf and FieldCompare.
    /// </summary>
    class ComparisonExpressionEvaluator : public IComparisonExpressionEvaluator
    {
    public:
        ComparisonExpressionEvaluator(SharedPtr<ComparisonEvaluationResult> result)
            : mInvocations(MakeObject<System::Collections::Generic::List<ArrayPtr<String>>>())
        {
            mResult = result;
        }

        SharedPtr<ComparisonEvaluationResult> Evaluate(SharedPtr<Field> field, SharedPtr<ComparisonExpression> expression) override
        {
            mInvocations->Add(MakeArray<String>({expression->get_LeftExpression(), expression->get_ComparisonOperator(), expression->get_RightExpression()}));

            return mResult;
        }

        SharedPtr<ExField::ComparisonExpressionEvaluator> AssertInvocationsCount(int expected)
        {
            EXPECT_EQ(expected, mInvocations->get_Count());
            return System::MakeSharedPtr(this);
        }

        SharedPtr<ExField::ComparisonExpressionEvaluator> AssertInvocationArguments(int invocationIndex, String expectedLeftExpression,
                                                                                    String expectedComparisonOperator, String expectedRightExpression)
        {
            ArrayPtr<String> arguments = mInvocations->idx_get(invocationIndex);

            EXPECT_EQ(expectedLeftExpression, arguments[0]);
            EXPECT_EQ(expectedComparisonOperator, arguments[1]);
            EXPECT_EQ(expectedRightExpression, arguments[2]);

            return System::MakeSharedPtr(this);
        }

    protected:
        virtual ~ComparisonExpressionEvaluator()
        {
        }

    private:
        SharedPtr<ComparisonEvaluationResult> mResult;
        SharedPtr<System::Collections::Generic::List<ArrayPtr<String>>> mInvocations;
    };
    //ExEnd

    void ComparisonExpressionEvaluatorNestedFields()
    {
        auto document = MakeObject<Document>();

        MakeObject<FieldBuilder>(FieldType::FieldIf)
            ->AddArgument(MakeObject<FieldBuilder>(FieldType::FieldIf)
                              ->AddArgument(123)
                              ->AddArgument(u">")
                              ->AddArgument(666)
                              ->AddArgument(u"left greater than right")
                              ->AddArgument(u"left less than right"))
            ->AddArgument(u"<>")
            ->AddArgument(MakeObject<FieldBuilder>(FieldType::FieldIf)
                              ->AddArgument(u"left expression")
                              ->AddArgument(u"=")
                              ->AddArgument(u"right expression")
                              ->AddArgument(u"expression are equal")
                              ->AddArgument(u"expression are not equal"))
            ->AddArgument(MakeObject<FieldBuilder>(FieldType::FieldIf)
                              ->AddArgument(MakeObject<FieldArgumentBuilder>()->AddText(u"#")->AddField(MakeObject<FieldBuilder>(FieldType::FieldPage)))
                              ->AddArgument(u"=")
                              ->AddArgument(MakeObject<FieldArgumentBuilder>()->AddText(u"#")->AddField(MakeObject<FieldBuilder>(FieldType::FieldNumPages)))
                              ->AddArgument(u"the last page")
                              ->AddArgument(u"not the last page"))
            ->AddArgument(MakeObject<FieldBuilder>(FieldType::FieldIf)
                              ->AddArgument(u"unexpected")
                              ->AddArgument(u"=")
                              ->AddArgument(u"unexpected")
                              ->AddArgument(u"unexpected")
                              ->AddArgument(u"unexpected"))
            ->BuildAndInsert(document->get_FirstSection()->get_Body()->get_FirstParagraph());

        auto evaluator = MakeObject<ExField::ComparisonExpressionEvaluator>(nullptr);
        document->get_FieldOptions()->set_ComparisonExpressionEvaluator(evaluator);

        document->UpdateFields();

        evaluator->AssertInvocationsCount(4)
            ->AssertInvocationArguments(0, u"123", u">", u"666")
            ->AssertInvocationArguments(1, u"\"left expression\"", u"=", u"\"right expression\"")
            ->AssertInvocationArguments(2, u"left less than right", u"<>", u"expression are not equal")
            ->AssertInvocationArguments(3, u"\"#1\"", u"=", u"\"#1\"");
    }

    void ComparisonExpressionEvaluatorHeaderFooterFields()
    {
        auto document = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(document);

        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::PageBreak);
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);

        MakeObject<FieldBuilder>(FieldType::FieldIf)
            ->AddArgument(MakeObject<FieldBuilder>(FieldType::FieldPage))
            ->AddArgument(u"=")
            ->AddArgument(MakeObject<FieldBuilder>(FieldType::FieldNumPages))
            ->AddArgument(MakeObject<FieldArgumentBuilder>()
                              ->AddField(MakeObject<FieldBuilder>(FieldType::FieldPage))
                              ->AddText(u" / ")
                              ->AddField(MakeObject<FieldBuilder>(FieldType::FieldNumPages)))
            ->AddArgument(MakeObject<FieldArgumentBuilder>()
                              ->AddField(MakeObject<FieldBuilder>(FieldType::FieldPage))
                              ->AddText(u" / ")
                              ->AddField(MakeObject<FieldBuilder>(FieldType::FieldNumPages)))
            ->BuildAndInsert(builder->get_CurrentParagraph());

        auto evaluator = MakeObject<ExField::ComparisonExpressionEvaluator>(nullptr);
        document->get_FieldOptions()->set_ComparisonExpressionEvaluator(evaluator);

        document->UpdateFields();

        evaluator->AssertInvocationsCount(3)
            ->AssertInvocationArguments(0, u"1", u"=", u"3")
            ->AssertInvocationArguments(1, u"2", u"=", u"3")
            ->AssertInvocationArguments(2, u"3", u"=", u"3");
    }

    //ExStart
    //ExFor:IFieldUpdatingCallback
    //ExFor:IFieldUpdatingCallback.FieldUpdating(Field)
    //ExFor:IFieldUpdatingCallback.FieldUpdated(Field)
    //ExSummary:Shows how to use callback methods during a field update.
    void FieldUpdatingCallbackTest()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u" DATE \\@ \"dddd, d MMMM yyyy\" ");
        builder->InsertField(u" TIME ");
        builder->InsertField(u" REVNUM ");
        builder->InsertField(u" AUTHOR  \"John Doe\" ");
        builder->InsertField(u" SUBJECT \"My Subject\" ");
        builder->InsertField(u" QUOTE \"Hello world!\" ");

        auto callback = MakeObject<ExField::FieldUpdatingCallback>();
        doc->get_FieldOptions()->set_FieldUpdatingCallback(callback);

        doc->UpdateFields();

        ASSERT_TRUE(callback->get_FieldUpdatedCalls()->Contains(u"Updating John Doe"));
    }

    /// <summary>
    /// Implement this interface if you want to have your own custom methods called during a field update.
    /// </summary>
    class FieldUpdatingCallback : public IFieldUpdatingCallback
    {
    public:
        SharedPtr<System::Collections::Generic::IList<String>> get_FieldUpdatedCalls()
        {
            return pr_FieldUpdatedCalls;
        }

        FieldUpdatingCallback()
        {
            pr_FieldUpdatedCalls = MakeObject<System::Collections::Generic::List<String>>();
        }

    private:
        SharedPtr<System::Collections::Generic::IList<String>> pr_FieldUpdatedCalls;

        /// <summary>
        /// A user defined method that is called just before a field is updated.
        /// </summary>
        void FieldUpdating(SharedPtr<Field> field) override
        {
            if (field->get_Type() == FieldType::FieldAuthor)
            {
                auto fieldAuthor = System::DynamicCast<FieldAuthor>(field);
                fieldAuthor->set_AuthorName(u"Updating John Doe");
            }
        }

        /// <summary>
        /// A user defined method that is called just after a field is updated.
        /// </summary>
        void FieldUpdated(SharedPtr<Field> field) override
        {
            get_FieldUpdatedCalls()->Add(field->get_Result());
        }
    };
    //ExEnd
};

} // namespace ApiExamples
