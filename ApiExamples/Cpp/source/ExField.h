#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Shading.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlock.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockBehavior.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockGallery.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/GlossaryDocument.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/VariableCollection.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldBuilder/FieldArgumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldBuilder/FieldBuilder.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldUpdateCultureSource.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldCreateDate.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldDate.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldEditTime.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldPrintDate.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldSaveDate.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldTime.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldCompare.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldDocVariable.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldGoToButton.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldIf.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldIfComparisonResult.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldMacroButton.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldPrint.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldAuthor.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldComments.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldDocProperty.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldFileSize.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldKeywords.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldLastSavedBy.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldNumChars.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldNumPages.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldNumWords.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldPrivate.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldSubject.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldTemplate.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldTitle.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/EquationsAndFormulas/FieldAdvance.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/EquationsAndFormulas/FieldFormula.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/EquationsAndFormulas/FieldSymbol.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/FormFields/FieldFormCheckBox.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/FormFields/FieldFormDropDown.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/FormFields/FieldFormText.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldIndex.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldRD.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldTA.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldTC.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldToa.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldToc.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldXE.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldAutoText.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldAutoTextList.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldBibliography.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldCitation.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldDde.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldDdeAuto.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldFillIn.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldGlossary.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldImport.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldInclude.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldIncludePicture.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldIncludeText.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldLink.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldNoteRef.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldPageRef.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldQuote.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldRef.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldSet.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldStyleRef.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldAddressBlock.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldMergeField.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldAutoNum.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldAutoNumLgl.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldAutoNumOut.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldListNum.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldPage.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldRevNum.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldSection.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldSectionPages.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldSeq.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Other/FieldAddIn.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Other/FieldData.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Other/FieldOcx.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Other/FieldUnknown.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Transitional/FieldBarcode.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Transitional/FieldBidiOutline.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Transitional/FieldDisplayBarcode.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Transitional/FieldEQ.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Transitional/FieldFootnoteRef.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Transitional/FieldInfo.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/UserInformation/FieldUserAddress.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/UserInformation/FieldUserInitials.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/UserInformation/FieldUserName.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/DropDownItemCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Format/FieldFormat.h>
#include <Aspose.Words.Cpp/Model/Fields/Format/GeneralFormat.h>
#include <Aspose.Words.Cpp/Model/Fields/Format/GeneralFormatCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldUserPromptRespondent.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldChar.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/UserInformation.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <drawing/color.h>
#include <drawing/color_translator.h>
#include <drawing/image.h>
#include <system/collections/dictionary.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ienumerator.h>
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
#include <system/scope_guard.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/test_tools.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/string_builder.h>
#include <system/threading/thread.h>
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

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;
using namespace Aspose::Words::Replacing;
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
        //ExSummary:Demonstrates how to retrieve the field class from an existing FieldStart node in the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto field = System::DynamicCast<FieldDate>(builder->InsertField(FieldType::FieldDate, true));
        field->get_Format()->set_DateTimeFormat(u"dddd, MMMM dd, yyyy");
        field->Update();

        SharedPtr<FieldChar> fieldStart = field->get_Start();
        ASSERT_EQ(FieldType::FieldDate, fieldStart->get_FieldType());
        ASPOSE_ASSERT_EQ(false, fieldStart->get_IsDirty());
        ASPOSE_ASSERT_EQ(false, fieldStart->get_IsLocked());

        // Retrieve the facade object which represents the field in the document
        field = System::DynamicCast<FieldDate>(fieldStart->GetField());

        ASPOSE_ASSERT_EQ(false, field->get_IsLocked());
        ASSERT_EQ(u" DATE  \\@ \"dddd, MMMM dd, yyyy\"", field->GetFieldCode());

        // Update the field to show the current date
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
        //ExSummary:Shows how to get text between field start and field separator (or field end if there is no separator).
        // Open a document which contains a MERGEFIELD inside an IF field
        auto doc = MakeObject<Document>(MyDir + u"Nested fields.docx");

        ASSERT_EQ(1, doc->get_Range()->get_Fields()->LINQ_Count([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldIf; }));
        //ExSkip

        // Get the outer IF field and print its full field code
        auto fieldIf = System::DynamicCast<FieldIf>(doc->get_Range()->get_Fields()->idx_get(0));
        std::cout << "Full field code including child fields:\n\t" << fieldIf->GetFieldCode() << std::endl;

        // All inner nested fields are printed by default
        ASSERT_EQ(fieldIf->GetFieldCode(), fieldIf->GetFieldCode(true));

        // Print the field code again but this time without the inner MERGEFIELD
        std::cout << "Field code with nested fields omitted:\n\t" << fieldIf->GetFieldCode(false) << std::endl;
        //ExEnd

        ASSERT_EQ(u" IF  > 0 \" (surplus of ) \" \"\" ", fieldIf->GetFieldCode(false));
        ASSERT_EQ(String::Format(u" IF {0} MERGEFIELD NetIncome {1}{2} > 0 \" (surplus of {3} MERGEFIELD  NetIncome \\f $ {4}{5}) \" \"\" ",
                                         ControlChar::FieldStartChar, ControlChar::FieldSeparatorChar, ControlChar::FieldEndChar, ControlChar::FieldStartChar,
                                         ControlChar::FieldSeparatorChar, ControlChar::FieldEndChar),
                  fieldIf->GetFieldCode(true));
    }

    void FieldDisplayResult()
    {
        //ExStart
        //ExFor:Field.DisplayResult
        //ExSummary:Shows how to get the text that represents the displayed field result.
        auto document = MakeObject<Document>(MyDir + u"Various fields.docx");

        SharedPtr<FieldCollection> fields = document->get_Range()->get_Fields();

        ASSERT_EQ(u"111", fields->idx_get(0)->get_DisplayResult());
        ASSERT_EQ(u"222", fields->idx_get(1)->get_DisplayResult());
        ASSERT_EQ(u"Multi\rLine\rText", fields->idx_get(2)->get_DisplayResult());
        ASSERT_EQ(u"%", fields->idx_get(3)->get_DisplayResult());
        ASSERT_EQ(u"Macro Button Text", fields->idx_get(4)->get_DisplayResult());
        ASSERT_EQ(String::Empty, fields->idx_get(5)->get_DisplayResult());

        // Method must be called to obtain correct value for the "FieldListNum", "FieldAutoNum",
        // "FieldAutoNumOut" and "FieldAutoNumLgl" fields
        document->UpdateListLabels();

        ASSERT_EQ(u"1)", fields->idx_get(5)->get_DisplayResult());
        //ExEnd
    }

    void CreateWithFieldBuilder()
    {
        //ExStart
        //ExFor:FieldBuilder.#ctor(FieldType)
        //ExFor:FieldBuilder.BuildAndInsert(Inline)
        //ExSummary:Builds and inserts a field into the document before the specified inline node.
        auto doc = MakeObject<Document>();

        // A convenient way of adding text content to a document is with a DocumentBuilder
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u" Hello world! This text is one Run, which is an inline node.");

        // Fields can be constructed in a similar way with a FieldBuilder, with arguments and switches added individually
        // In this case we will construct a BARCODE field which represents a US postal code
        auto fieldBuilder = MakeObject<FieldBuilder>(FieldType::FieldBarcode);
        fieldBuilder->AddArgument(u"90210");
        fieldBuilder->AddSwitch(u"\\f", u"A");
        fieldBuilder->AddSwitch(u"\\u");

        // Insert the field before any inline node
        fieldBuilder->BuildAndInsert(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0));
        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.CreateWithFieldBuilder.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.CreateWithFieldBuilder.docx");

        TestUtil::VerifyField(FieldType::FieldBarcode, u" BARCODE 90210 \\f A \\u ", String::Empty, doc->get_Range()->get_Fields()->idx_get(0));

        ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(11)->get_PreviousSibling(),
                         doc->get_Range()->get_Fields()->idx_get(0)->get_End());
        ASSERT_EQ(String::Format(u"{0} BARCODE 90210 \\f A \\u {1} Hello world! This text is one Run, which is an inline node.",
                                         ControlChar::FieldStartChar, ControlChar::FieldEndChar),
                  doc->GetText().Trim());
    }

    void CreateRevNumFieldByDocumentBuilder()
    {
        //ExStart
        //ExFor:FieldRevNum
        //ExSummary:Shows how to work with REVNUM fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add some text to a blank document with a DocumentBuilder
        builder->Write(u"Current revision #");

        // Insert a REVNUM field, which displays the document's current revision number property
        auto field = System::DynamicCast<FieldRevNum>(builder->InsertField(FieldType::FieldRevisionNum, true));

        ASSERT_EQ(u" REVNUM ", field->GetFieldCode());
        ASSERT_EQ(u"1", field->get_Result());
        ASSERT_EQ(1, doc->get_BuiltInDocumentProperties()->get_RevisionNumber());

        // This property counts how many times a document has been saved in Microsoft Word, is unrelated to revision tracking,
        // can be found by right clicking the document in Windows Explorer via Properties > Details
        // This property is only manually updated by Aspose.Words
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

    void CreateInfoFieldWithFieldBuilder()
    {
        auto doc = MakeObject<Document>();
        SharedPtr<Run> run = DocumentHelper::InsertNewRun(doc, u" Hello World!", 0);

        auto fieldBuilder = MakeObject<FieldBuilder>(FieldType::FieldInfo);
        fieldBuilder->BuildAndInsert(run);

        doc->UpdateFields();
        doc = DocumentHelper::SaveOpen(doc);

        auto info = System::DynamicCast<FieldInfo>(doc->get_Range()->get_Fields()->idx_get(0));
        ASSERT_FALSE(info == nullptr);
    }

    void CreateInfoFieldWithDocumentBuilder()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u"INFO MERGEFORMAT");

        doc = DocumentHelper::SaveOpen(doc);

        auto info = System::DynamicCast<FieldInfo>(doc->get_Range()->get_Fields()->idx_get(0));
        ASSERT_FALSE(info == nullptr);
    }

    void GetFieldFromFieldCollection()
    {
        auto doc = MakeObject<Document>(MyDir + u"Table of contents.docx");

        SharedPtr<Field> field = doc->get_Range()->get_Fields()->idx_get(0);

        // This should be the first field in the document - a TOC field
        std::cout << System::EnumGetName(field->get_Type()) << std::endl;
    }

    void InsertFieldNone()
    {
        //ExStart
        //ExFor:FieldUnknown
        //ExSummary:Shows how to work with 'FieldNone' field in a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a field that does not denote a real field type in its field code
        SharedPtr<Field> field = builder->InsertField(u" NOTAREALFIELD //a");

        // Fields like that can be written and read, and are assigned a special "FieldNone" type
        ASSERT_EQ(FieldType::FieldNone, field->get_Type());

        // We can also still work with these fields, and assign them as instances of a special "FieldUnknown" class
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

        // Insert a TC field at the current document builder position
        builder->InsertField(u"TC \"Entry Text\" \\f t");
    }

    void FieldLocale()
    {
        //ExStart
        //ExFor:Field.LocaleId
        //ExSummary:Shows how to insert a field and work with its locale.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a DATE field and print the date it will display, formatted according to your thread's current culture
        SharedPtr<Field> field = builder->InsertField(u"DATE");
        std::cout << "Today's date, as displayed in the \"" << System::Globalization::CultureInfo::get_CurrentCulture()->get_EnglishName()
                  << "\" culture: " << field->get_Result() << std::endl;

        ASSERT_EQ(1033, field->get_LocaleId());
        ASSERT_EQ(FieldUpdateCultureSource::CurrentThread, doc->get_FieldOptions()->get_FieldUpdateCultureSource());
        //ExSkip

        // We can get the field to display a date in a different format if we change the current thread's culture
        // If we want to avoid making such an all-encompassing change,
        // we can set this option to get the document's fields to get their culture from themselves
        // Then, we can change a field's LocaleId and it will display its result in any culture we choose
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

    void ChangeLocale()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u"MERGEFIELD Date");

        // Store the current culture so it can be set back once mail merge is complete
        SharedPtr<System::Globalization::CultureInfo> currentCulture = System::Threading::Thread::get_CurrentThread()->get_CurrentCulture();
        // Set to German language so dates and numbers are formatted using this culture during mail merge
        System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(MakeObject<System::Globalization::CultureInfo>(u"de-DE"));

        doc->get_MailMerge()->Execute(
            MakeArray<String>({u"Date"}),
            MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<System::DateTime>(System::DateTime::get_Now())}));

        // Restore the original culture and save the document
        System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(currentCulture);
        doc->Save(ArtifactsDir + u"Field.ChangeLocale.docx");
    }

    void RemoveTocFromDocument()
    {
        // Open a document which contains a TOC
        auto doc = MakeObject<Document>(MyDir + u"Table of contents.docx");

        // Remove the first TOC from the document
        SharedPtr<Field> tocField = doc->get_Range()->get_Fields()->idx_get(0);
        tocField->Remove();

        doc->Save(ArtifactsDir + u"Field.RemoveTocFromDocument.docx");
    }

    void InsertTcFieldsAtText()
    {
        auto doc = MakeObject<Document>();

        auto options = MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(MakeObject<ExField::InsertTcFieldHandler>(u"Chapter 1", u"\\l 1"));

        // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document
        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"The Beginning"), u"", options);
    }

private:
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

        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            // Create a builder to insert the field
            auto builder = MakeObject<DocumentBuilder>(System::DynamicCast<Document>(args->get_MatchNode()->get_Document()));
            // Move to the first node of the match
            builder->MoveTo(args->get_MatchNode());

            // If the user specified text to be used in the field as display text then use that, otherwise use the
            // match String as the display text
            String insertText = !String::IsNullOrEmpty(mFieldText) ? mFieldText : args->get_Match()->get_Value();

            // Insert the TC field before this node using the specified String as the display text and user defined switches
            builder->InsertField(String::Format(u"TC \"{0}\" {1}", insertText, mFieldSwitches));

            // We have done what we want so skip replacement
            return ReplaceAction::Skip;
        }

    private:
        String mFieldText;
        String mFieldSwitches;
    };

public:
    void UpdateDirtyFields(bool doUpdateDirtyFields)
    {
        //ExStart
        //ExFor:Field.IsDirty
        //ExFor:LoadOptions.UpdateDirtyFields
        //ExSummary:Shows how to use special property for updating field result.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Give the document's built in property "Author" a value and display it with a field
        doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
        auto field = System::DynamicCast<FieldAuthor>(builder->InsertField(FieldType::FieldAuthor, true));

        ASSERT_FALSE(field->get_IsDirty());
        ASSERT_EQ(u"John Doe", field->get_Result());

        // Update the "Author" property
        doc->get_BuiltInDocumentProperties()->set_Author(u"John & Jane Doe");

        // Fields of some types, such as AUTHOR, do not update according to their source values in real time,
        // and need to be updated manually beforehand every time an accurate value is required
        ASSERT_EQ(u"John Doe", field->get_Result());

        // Since the field's value is out of date, we can mark it as "Dirty"
        field->set_IsDirty(true);

        {
            auto docStream = MakeObject<System::IO::MemoryStream>();
            doc->Save(docStream, SaveFormat::Docx);

            // Re-open the document from the stream while using a LoadOptions object to specify
            // whether to update all fields marked as "Dirty" in the process, so they can display accurate values immediately
            auto options = MakeObject<LoadOptions>();
            options->set_UpdateDirtyFields(doUpdateDirtyFields);
            doc = MakeObject<Document>(docStream, options);

            ASSERT_EQ(u"John & Jane Doe", doc->get_BuiltInDocumentProperties()->get_Author());

            field = System::DynamicCast<FieldAuthor>(doc->get_Range()->get_Fields()->idx_get(0));

            if (doUpdateDirtyFields)
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

        // Add some text into the paragraph
        SharedPtr<Run> run = DocumentHelper::InsertNewRun(doc, u" Hello World!", 0);

        auto argumentBuilder = MakeObject<FieldArgumentBuilder>();
        argumentBuilder->AddField(MakeObject<FieldBuilder>(FieldType::FieldMergeField));
        argumentBuilder->AddNode(run);
        argumentBuilder->AddText(u"Text argument builder");

        auto fieldBuilder = MakeObject<FieldBuilder>(FieldType::FieldIncludeText);

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<SharedPtr<Field>()> _local_func_1 = [&fieldBuilder, &argumentBuilder, &run]()
        {
            return fieldBuilder->AddArgument(argumentBuilder)
                ->AddArgument(u"=")
                ->AddArgument(u"BestField")
                ->AddArgument(10)
                ->AddArgument(20.0)
                ->BuildAndInsert(run);
        };

        ASSERT_THROW(_local_func_1(), System::ArgumentException);
    }

    void UpdateFieldIgnoringMergeFormat()
    {
        //ExStart
        //ExFor:Field.Update(bool)
        //ExFor:LoadOptions.PreserveIncludePictureField
        //ExSummary:Shows a way to update a field ignoring the MERGEFORMAT switch.
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_PreserveIncludePictureField(true);
        auto doc = MakeObject<Document>(MyDir + u"Field sample - INCLUDEPICTURE.docx", loadOptions);

        auto includePicture = System::DynamicCast<FieldIncludePicture>(doc->get_Range()->get_Fields()->LINQ_First([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldIncludePicture; }));
        includePicture->set_SourceFullName(ImageDir + u"Transparent background logo.png");
        includePicture->Update(true);

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.UpdateFieldIgnoringMergeFormat.docx");
        //ExEnd

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Aspose::Words::Fields::Field> f)> _local_func_3 = [](SharedPtr<Aspose::Words::Fields::Field> f)
        {
            return f->get_Type() == FieldType::FieldIncludePicture;
        };

        ASSERT_TRUE(doc->get_Range()->get_Fields()->LINQ_Any(static_cast<System::Func<SharedPtr<Field>, bool>>(_local_func_3)));

        doc = MakeObject<Document>(ArtifactsDir + u"Field.UpdateFieldIgnoringMergeFormat.docx");
        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_TRUE(shape->get_IsImage());

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Aspose::Words::Fields::Field> f)> _local_func_4 = [](SharedPtr<Aspose::Words::Fields::Field> f)
        {
            return f->get_Type() == FieldType::FieldIncludePicture;
        };

        ASSERT_FALSE(doc->get_Range()->get_Fields()->LINQ_Any(static_cast<System::Func<SharedPtr<Field>, bool>>(_local_func_4)));
    }

    void FieldFormat_()
    {
        //ExStart
        //ExFor:Field.Format
        //ExFor:Field.Update
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
        //ExSummary:Shows how to format fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to insert field with no format
        SharedPtr<Field> field = builder->InsertField(u"= 2 + 3");

        ASSERT_EQ(u"5", field->get_Result());

        // We can format our field here instead of in the field code
        SharedPtr<FieldFormat> format = field->get_Format();
        format->set_NumericFormat(u"$###.00");

        // Fields need to be manually updated in order for formatting to take effect
        field->Update();

        ASSERT_EQ(u"$  5.00", field->get_Result());

        // Apply a date/time format
        field = builder->InsertField(u"DATE");
        format = field->get_Format();
        format->set_DateTimeFormat(u"dddd, MMMM dd, yyyy");
        field->Update();

        std::cout << "Today's date, in " << format->get_DateTimeFormat() << " format:\n\t" << field->get_Result() << std::endl;

        // Apply 2 general formats at the same time
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

        ASSERT_EQ(u"LVIII", field->get_Result());
        ASSERT_EQ(2, format->get_GeneralFormats()->get_Count());
        ASSERT_EQ(GeneralFormat::LowercaseRoman, format->get_GeneralFormats()->idx_get(0));

        // Removing field formats
        format->get_GeneralFormats()->Remove(GeneralFormat::LowercaseRoman);
        format->get_GeneralFormats()->RemoveAt(0);
        ASSERT_EQ(0, format->get_GeneralFormats()->get_Count());
        field->Update();

        // Our field has no general formats left and is back to default form
        ASSERT_EQ(u"58", field->get_Result());
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);

        ASSERT_EQ(u"$###.00", doc->get_Range()->get_Fields()->idx_get(0)->get_Format()->get_NumericFormat());
        ASSERT_EQ(u"$  5.00", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());

        ASSERT_EQ(u"dddd, MMMM dd, yyyy", doc->get_Range()->get_Fields()->idx_get(1)->get_Format()->get_DateTimeFormat());
        ASSERT_EQ(System::DateTime::get_Today(), System::DateTime::Parse(doc->get_Range()->get_Fields()->idx_get(1)->get_Result()));

        ASSERT_EQ(0, doc->get_Range()->get_Fields()->idx_get(2)->get_Format()->get_GeneralFormats()->get_Count());
        ASSERT_EQ(u"58", doc->get_Range()->get_Fields()->idx_get(2)->get_Result());
    }

    void UnlinkAllFieldsInDocument()
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
        //ExSummary:Shows how to unlink all fields in range.
        auto doc = MakeObject<Document>(MyDir + u"Linked fields.docx");

        auto newSection = System::DynamicCast<Section>(System::StaticCast<Node>(doc->get_Sections()->idx_get(0))->Clone(true));
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
        //ExSummary:Shows how to unlink specific field.
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

        for (auto para : System::IterateOver(paragraphCollection->LINQ_OfType<SharedPtr<Paragraph>>()))
        {
            // Check all runs in the paragraph for the first page breaks
            for (auto run : System::IterateOver(para->get_Runs()->LINQ_OfType<SharedPtr<Run>>()))
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

        for (auto field : System::IterateOver(fStart->LINQ_OfType<SharedPtr<FieldStart>>()))
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

protected:
    static void RemoveSequence(SharedPtr<Node> start, SharedPtr<Node> end)
    {
        SharedPtr<Node> curNode = start->NextPreOrder(start->get_Document());
        while (curNode != nullptr && !System::ObjectExt::Equals(curNode, end))
        {
            // Move to next node
            SharedPtr<Node> nextNode = curNode->NextPreOrder(start->get_Document());

            // Check whether current contains end node
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

public:
    void DropDownItemCollection_()
    {
        //ExStart
        //ExFor:Fields.DropDownItemCollection
        //ExFor:Fields.DropDownItemCollection.Add(String)
        //ExFor:Fields.DropDownItemCollection.Clear
        //ExFor:Fields.DropDownItemCollection.Contains(String)
        //ExFor:Fields.DropDownItemCollection.Count
        //ExFor:Fields.DropDownItemCollection.GetEnumerator
        //ExFor:Fields.DropDownItemCollection.IndexOf(String)
        //ExFor:Fields.DropDownItemCollection.Insert(Int32, String)
        //ExFor:Fields.DropDownItemCollection.Item(Int32)
        //ExFor:Fields.DropDownItemCollection.Remove(String)
        //ExFor:Fields.DropDownItemCollection.RemoveAt(Int32)
        //ExSummary:Shows how to insert a combo box field and manipulate the elements in its item collection.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to create and populate a combo box
        ArrayPtr<String> items = MakeArray<String>({u"One", u"Two", u"Three"});
        SharedPtr<FormField> comboBoxField = builder->InsertComboBox(u"DropDown", items, 0);

        // Get the list of drop-down items
        SharedPtr<DropDownItemCollection> dropDownItems = comboBoxField->get_DropDownItems();

        ASSERT_EQ(3, dropDownItems->get_Count());
        ASSERT_EQ(u"One", dropDownItems->idx_get(0));
        ASSERT_EQ(1, dropDownItems->IndexOf(u"Two"));
        ASSERT_TRUE(dropDownItems->Contains(u"Three"));

        // We can add an item to the end of the collection or insert it at a desired index
        dropDownItems->Add(u"Four");
        dropDownItems->Insert(3, u"Three and a half");
        ASSERT_EQ(5, dropDownItems->get_Count());

        // Iterate over the collection and print every element
        {
            SharedPtr<System::Collections::Generic::IEnumerator<String>> dropDownCollectionEnumerator = dropDownItems->GetEnumerator();
            while (dropDownCollectionEnumerator->MoveNext())
            {
                std::cout << dropDownCollectionEnumerator->get_Current() << std::endl;
            }
        }

        // We can remove elements in the same way we added them
        dropDownItems->Remove(u"Four");
        dropDownItems->RemoveAt(3);
        ASSERT_FALSE(dropDownItems->Contains(u"Three and a half"));
        ASSERT_FALSE(dropDownItems->Contains(u"Four"));

        doc->Save(ArtifactsDir + u"Field.DropDownItemCollection.docx");

        // Empty the collection
        dropDownItems->Clear();
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        dropDownItems = doc->get_Range()->get_FormFields()->idx_get(0)->get_DropDownItems();

        ASSERT_EQ(0, dropDownItems->get_Count());

        doc = MakeObject<Document>(ArtifactsDir + u"Field.DropDownItemCollection.docx");
        dropDownItems = doc->get_Range()->get_FormFields()->idx_get(0)->get_DropDownItems();

        ASSERT_EQ(3, dropDownItems->get_Count());
        ASSERT_EQ(u"One", dropDownItems->idx_get(0));
        ASSERT_EQ(u"Two", dropDownItems->idx_get(1));
        ASSERT_EQ(u"Three", dropDownItems->idx_get(2));
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
        //ExSummary:Shows how to insert an advance field and edit its properties.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"This text is in its normal place.");
        // Create an advance field using document builder
        auto field = System::DynamicCast<FieldAdvance>(builder->InsertField(FieldType::FieldAdvance, true));

        builder->Write(u"This text is moved up and to the right.");

        ASSERT_EQ(FieldType::FieldAdvance, field->get_Type());
        //ExSkip
        ASSERT_EQ(u" ADVANCE ", field->GetFieldCode());
        //ExSkip
        // The second text that the builder added will now be moved
        field->set_RightOffset(u"5");
        field->set_UpOffset(u"5");

        ASSERT_EQ(u" ADVANCE  \\r 5 \\u 5", field->GetFieldCode());
        // If we want to move text in the other direction, and try do that by using negative values for the above field members, we will get an error in our
        // document Instead, we need to specify a positive value for the opposite respective field directional variable
        field = System::DynamicCast<FieldAdvance>(builder->InsertField(FieldType::FieldAdvance, true));
        field->set_DownOffset(u"5");
        field->set_LeftOffset(u"100");

        ASSERT_EQ(u" ADVANCE  \\d 5 \\l 100", field->GetFieldCode());
        // We are still on one paragraph
        ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->get_Paragraphs()->get_Count());
        // Since we're setting horizontal and vertical positions next, ending the paragraph prevents the previous line from getting moved with the next one
        builder->Writeln(u"This text is moved down and to the left, overlapping the previous text.");
        // This time we can also use negative values
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
        //ExSummary:Shows how to build a field address block.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to insert a field address block
        auto field = System::DynamicCast<FieldAddressBlock>(builder->InsertField(FieldType::FieldAddressBlock, true));

        // Initially our field is an empty address block field with null attributes
        ASSERT_EQ(u" ADDRESSBLOCK ", field->GetFieldCode());

        // Setting this to "2" will cause all countries/regions to be included, unless it is the one specified in the ExcludedCountryOrRegionName attribute
        field->set_IncludeCountryOrRegionName(u"2");
        field->set_FormatAddressOnCountryOrRegion(true);
        field->set_ExcludedCountryOrRegionName(u"United States");

        // Specify our own name and address format
        field->set_NameAndAddressFormat(u"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>");

        // By default, the language ID will be set to that of the first character of the document
        // In this case we will specify it to be English
        field->set_LanguageId(u"1033");

        // Our field code has changed according to the attribute values that we set
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
    //ExFor:FieldCollection.Clear
    //ExFor:FieldCollection.Count
    //ExFor:FieldCollection.GetEnumerator
    //ExFor:FieldCollection.Item(Int32)
    //ExFor:FieldCollection.Remove(Field)
    //ExFor:FieldCollection.Remove(FieldStart)
    //ExFor:FieldCollection.RemoveAt(Int32)
    //ExFor:FieldStart
    //ExFor:FieldStart.Accept(DocumentVisitor)
    //ExFor:FieldSeparator
    //ExFor:FieldSeparator.Accept(DocumentVisitor)
    //ExFor:FieldEnd
    //ExFor:FieldEnd.Accept(DocumentVisitor)
    //ExFor:FieldEnd.HasSeparator
    //ExFor:Field.End
    //ExFor:Field.Remove
    //ExFor:Field.Separator
    //ExFor:Field.Start
    //ExSummary:Shows how to work with a document's field collection.
    void FieldCollection_()
    {
        // Create a new document and insert some fields
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u" DATE \\@ \"dddd, d MMMM yyyy\" ");
        builder->InsertField(u" TIME ");
        builder->InsertField(u" REVNUM ");
        builder->InsertField(u" AUTHOR  \"John Doe\" ");
        builder->InsertField(u" SUBJECT \"My Subject\" ");
        builder->InsertField(u" QUOTE \"Hello world!\" ");
        doc->UpdateFields();

        // Get the collection that contains all the fields in a document
        SharedPtr<FieldCollection> fields = doc->get_Range()->get_Fields();
        ASSERT_EQ(6, fields->get_Count());

        // Iterate over the field collection and print contents and type of every field using a custom visitor implementation
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

        // Get a field to remove itself
        fields->idx_get(0)->Remove();
        ASSERT_EQ(5, fields->get_Count());

        // Remove a field by reference
        SharedPtr<Field> lastField = fields->idx_get(3);
        fields->Remove(lastField);
        ASSERT_EQ(4, fields->get_Count());

        // Remove a field by index
        fields->RemoveAt(2);
        ASSERT_EQ(3, fields->get_Count());

        // Remove all fields from the document
        fields->Clear();
        ASSERT_EQ(0, fields->get_Count());
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

    void FieldCompare_()
    {
        //ExStart
        //ExFor:FieldCompare
        //ExFor:FieldCompare.ComparisonOperator
        //ExFor:FieldCompare.LeftExpression
        //ExFor:FieldCompare.RightExpression
        //ExSummary:Shows how to insert a field that compares expressions.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a compare field using a document builder
        auto field = System::DynamicCast<FieldCompare>(builder->InsertField(FieldType::FieldCompare, true));

        // Construct a comparison statement
        field->set_LeftExpression(u"3");
        field->set_ComparisonOperator(u"<");
        field->set_RightExpression(u"2");

        // The compare field will print a "0" or "1" depending on the truth of its statement
        // The result of this statement is false, so a "0" will be show up in the document
        ASSERT_EQ(u" COMPARE  3 < 2", field->GetFieldCode());

        builder->Writeln();

        // Here a "1" will show up, because the statement is true
        field = System::DynamicCast<FieldCompare>(builder->InsertField(FieldType::FieldCompare, true));
        field->set_LeftExpression(u"5");
        field->set_ComparisonOperator(u"=");
        field->set_RightExpression(u"2 + 3");

        ASSERT_EQ(u" COMPARE  5 = \"2 + 3\"", field->GetFieldCode());

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
        //ExSummary:Shows how to insert an if field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Statement 1: ");

        // Use document builder to insert an if field
        auto field = System::DynamicCast<FieldIf>(builder->InsertField(FieldType::FieldIf, true));

        // The if field will output either the TrueText or FalseText string into the document, depending on the truth of the statement
        // In this case, "0 = 1" is incorrect, so the output will be "False"
        field->set_LeftExpression(u"0");
        field->set_ComparisonOperator(u"=");
        field->set_RightExpression(u"1");
        field->set_TrueText(u"True");
        field->set_FalseText(u"False");

        ASSERT_EQ(u" IF  0 = 1 True False", field->GetFieldCode());
        ASSERT_EQ(FieldIfComparisonResult::False, field->EvaluateCondition());

        // This time, the statement is correct, so the output will be "True"
        builder->Write(u"\nStatement 2: ");
        field = System::DynamicCast<FieldIf>(builder->InsertField(FieldType::FieldIf, true));
        field->set_LeftExpression(u"5");
        field->set_ComparisonOperator(u"=");
        field->set_RightExpression(u"2 + 3");
        field->set_TrueText(u"True");
        field->set_FalseText(u"False");

        ASSERT_EQ(u" IF  5 = \"2 + 3\" True False", field->GetFieldCode());
        ASSERT_EQ(FieldIfComparisonResult::True, field->EvaluateCondition());

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

        // The two fields we insert here will be automatically numbered 1 and 2
        auto field = System::DynamicCast<FieldAutoNum>(builder->InsertField(FieldType::FieldAutoNum, true));
        builder->Writeln(u"\tParagraph 1.");

        ASSERT_EQ(u" AUTONUM ", field->GetFieldCode());

        field = System::DynamicCast<FieldAutoNum>(builder->InsertField(FieldType::FieldAutoNum, true));
        builder->Writeln(u"\tParagraph 2.");

        // Leaving the FieldAutoNum.SeparatorCharacter field null will set the separator character to '.' by default
        ASSERT_TRUE(field->get_SeparatorCharacter() == nullptr);

        // The first character of the string entered here will be used as the separator character
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

        // Set a filler paragraph string
        const String loremIpsum =
            String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ") +
            u"\nUt enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ";

        // In this case our AUTONUMLGL field will number our first paragraph as "1."
        InsertNumberedClause(builder, u"\tHeading 1", loremIpsum, StyleIdentifier::Heading1);

        // Our heading style number will be 1 again, so this field will keep counting headings at a heading level of 1
        InsertNumberedClause(builder, u"\tHeading 2", loremIpsum, StyleIdentifier::Heading1);

        // Our heading style is 2, setting the paragraph numbering depth to 2, setting this field's value to "2.1."
        InsertNumberedClause(builder, u"\tHeading 3", loremIpsum, StyleIdentifier::Heading2);

        // Our heading style is 3, so we are going deeper again to "2.1.1."
        InsertNumberedClause(builder, u"\tHeading 4", loremIpsum, StyleIdentifier::Heading3);

        // Our heading style is 2, and the next field number at that level is "2.2."
        InsertNumberedClause(builder, u"\tHeading 5", loremIpsum, StyleIdentifier::Heading2);

        auto isAutoNumLegal = [](SharedPtr<Field> f)
        {
            return f->get_Type() == FieldType::FieldAutoNumLegal;
        };

        for (auto field : System::IterateOver<FieldAutoNumLgl>(doc->get_Range()->get_Fields()->LINQ_Where(isAutoNumLegal)))
        {
            // By default, the separator will appear as "." in the document but here it is null
            ASSERT_TRUE(field->get_SeparatorCharacter() == nullptr);

            // Change the separator character and remove trailing separators
            field->set_SeparatorCharacter(u":");
            field->set_RemoveTrailingPeriod(true);
            ASSERT_EQ(u" AUTONUMLGL  \\s : \\e", field->GetFieldCode());
        }

        doc->Save(ArtifactsDir + u"Field.AUTONUMLGL.docx");
        TestFieldAutoNumLgl(doc);
        //ExSkip
    }

protected:
    /// <summary>
    /// Get a document builder to insert a clause numbered by an AUTONUMLGL field.
    /// </summary>
    static void InsertNumberedClause(SharedPtr<DocumentBuilder> builder, String heading, String contents, StyleIdentifier headingStyle)
    {
        // This legal field will automatically number our clauses, taking heading style level into account
        builder->InsertField(FieldType::FieldAutoNumLegal, true);
        builder->get_CurrentParagraph()->get_ParagraphFormat()->set_StyleIdentifier(headingStyle);
        builder->Writeln(heading);

        // This text will belong to the auto num legal field above it
        // It will collapse when the arrow next to the corresponding AUTONUMLGL field is clicked in MS Word
        builder->get_CurrentParagraph()->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::BodyText);
        builder->Writeln(contents);
    }
    //ExEnd

    void TestFieldAutoNumLgl(SharedPtr<Document> doc)
    {
        doc = DocumentHelper::SaveOpen(doc);

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Aspose::Words::Fields::Field> f)> _local_func_6 = [](SharedPtr<Aspose::Words::Fields::Field> f)
        {
            return f->get_Type() == FieldType::FieldAutoNumLegal;
        };

        for (auto field : System::IterateOver<FieldAutoNumLgl>(
                 doc->get_Range()->get_Fields()->LINQ_Where(static_cast<System::Func<SharedPtr<Field>, bool>>(_local_func_6))))
        {
            TestUtil::VerifyField(FieldType::FieldAutoNumLegal, u" AUTONUMLGL  \\s : \\e", String::Empty, field);

            ASSERT_EQ(u":", field->get_SeparatorCharacter());
            ASSERT_TRUE(field->get_RemoveTrailingPeriod());
        }
    }

public:
    void FieldAutoNumOut_()
    {
        //ExStart
        //ExFor:FieldAutoNumOut
        //ExSummary:Shows how to number paragraphs using AUTONUMOUT fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // The two fields that we insert here will be numbered 1 and 2
        builder->InsertField(FieldType::FieldAutoNumOutline, true);
        builder->Writeln(u"\tParagraph 1.");
        builder->InsertField(FieldType::FieldAutoNumOutline, true);
        builder->Writeln(u"\tParagraph 2.");

        for (auto field : System::IterateOver<FieldAutoNumOut>(doc->get_Range()->get_Fields()->LINQ_Where([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldAutoNumOutline; })))
        {
            ASSERT_EQ(u" AUTONUMOUT ", field->GetFieldCode());
        }

        doc->Save(ArtifactsDir + u"Field.AUTONUMOUT.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.AUTONUMOUT.docx");

        for (auto field : System::IterateOver(doc->get_Range()->get_Fields()))
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
        //ExSummary:Shows how to insert a building block into a document and display it with AUTOTEXT and GLOSSARY fields.
        auto doc = MakeObject<Document>();

        // Create a glossary document and add an AutoText building block
        doc->set_GlossaryDocument(MakeObject<GlossaryDocument>());
        auto buildingBlock = MakeObject<BuildingBlock>(doc->get_GlossaryDocument());
        buildingBlock->set_Name(u"MyBlock");
        buildingBlock->set_Gallery(BuildingBlockGallery::AutoText);
        buildingBlock->set_Category(u"General");
        buildingBlock->set_Description(u"MyBlock description");
        buildingBlock->set_Behavior(BuildingBlockBehavior::Paragraph);
        doc->get_GlossaryDocument()->AppendChild(buildingBlock);

        // Create a source and add it as text content to our building block
        auto buildingBlockSource = MakeObject<Document>();
        auto buildingBlockSourceBuilder = MakeObject<DocumentBuilder>(buildingBlockSource);
        buildingBlockSourceBuilder->Writeln(u"Hello World!");

        SharedPtr<Node> buildingBlockContent = doc->get_GlossaryDocument()->ImportNode(buildingBlockSource->get_FirstSection(), true);
        buildingBlock->AppendChild(buildingBlockContent);

        // Create an advance field using document builder
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto fieldAutoText = System::DynamicCast<FieldAutoText>(builder->InsertField(FieldType::FieldAutoText, true));

        // Refer to our building block by name
        fieldAutoText->set_EntryName(u"MyBlock");

        ASSERT_EQ(u" AUTOTEXT  MyBlock", fieldAutoText->GetFieldCode());

        // Put additional templates here
        doc->get_FieldOptions()->set_BuiltInTemplatesPaths(MakeArray<String>({MyDir + u"Busniess brochure.dotx"}));

        // We can also display our building block with a GLOSSARY field
        auto fieldGlossary = System::DynamicCast<FieldGlossary>(builder->InsertField(FieldType::FieldGlossary, true));
        fieldGlossary->set_EntryName(u"MyBlock");

        ASSERT_EQ(u" GLOSSARY  MyBlock", fieldGlossary->GetFieldCode());

        // The text content of our building block will be visible in the output
        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.AUTOTEXT.dotx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.AUTOTEXT.dotx");

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
    //ExSummary:Shows how to use an AutoTextList field to select from a list of AutoText entries.
    void FieldAutoTextList_()
    {
        auto doc = MakeObject<Document>();

        // Create a glossary document and populate it with auto text entries that our auto text list will let us select from
        doc->set_GlossaryDocument(MakeObject<GlossaryDocument>());
        AppendAutoTextEntry(doc->get_GlossaryDocument(), u"AutoText 1", u"Contents of AutoText 1");
        AppendAutoTextEntry(doc->get_GlossaryDocument(), u"AutoText 2", u"Contents of AutoText 2");
        AppendAutoTextEntry(doc->get_GlossaryDocument(), u"AutoText 3", u"Contents of AutoText 3");

        // Insert an auto text list using a document builder and change its properties
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto field = System::DynamicCast<FieldAutoTextList>(builder->InsertField(FieldType::FieldAutoTextList, true));
        // This is the text that will be visible in the document
        field->set_EntryName(u"Right click here to pick an AutoText block");
        field->set_ListStyle(u"Heading 1");
        field->set_ScreenTip(u"Hover tip text for AutoTextList goes here");

        ASSERT_EQ(String(u" AUTOTEXTLIST  \"Right click here to pick an AutoText block\" ") + u"\\s \"Heading 1\" " +
                      u"\\t \"Hover tip text for AutoTextList goes here\"",
                  field->GetFieldCode());

        doc->Save(ArtifactsDir + u"Field.AUTOTEXTLIST.dotx");
    }

protected:
    /// <summary>
    /// Create an AutoText entry and add it to a glossary document.
    /// </summary>
    static void AppendAutoTextEntry(SharedPtr<GlossaryDocument> glossaryDoc, String name, String contents)
    {
        // Create building block and set it up as an auto text entry
        auto buildingBlock = MakeObject<BuildingBlock>(glossaryDoc);
        buildingBlock->set_Name(name);
        buildingBlock->set_Gallery(BuildingBlockGallery::AutoText);
        buildingBlock->set_Category(u"General");
        buildingBlock->set_Behavior(BuildingBlockBehavior::Paragraph);

        // Add content to the building block
        auto section = MakeObject<Section>(glossaryDoc);
        section->AppendChild(MakeObject<Body>(glossaryDoc));
        section->get_Body()->AppendParagraph(contents);
        buildingBlock->AppendChild(section);

        // Add auto text entry to glossary document
        glossaryDoc->AppendChild(buildingBlock);
    }
    //ExEnd

public:
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

        // Insert a list num field using a document builder
        auto field = System::DynamicCast<FieldListNum>(builder->InsertField(FieldType::FieldListNum, true));

        // Lists start counting at 1 by default, but we can change this number at any time
        // In this case, we will do a zero-based count
        field->set_StartingNumber(u"0");
        builder->Writeln(u"Paragraph 1");

        ASSERT_EQ(u" LISTNUM  \\s 0", field->GetFieldCode());

        // Placing several list num fields in one paragraph increases the list level instead of the current number,
        // in this case resulting in "1)a)i)", list level 3
        builder->InsertField(FieldType::FieldListNum, true);
        builder->InsertField(FieldType::FieldListNum, true);
        builder->InsertField(FieldType::FieldListNum, true);
        builder->Writeln(u"Paragraph 2");

        // The list level resets with new paragraphs, so to keep counting at a desired list level, we need to set the ListLevel property accordingly
        field = System::DynamicCast<FieldListNum>(builder->InsertField(FieldType::FieldListNum, true));
        field->set_ListLevel(u"3");
        builder->Writeln(u"Paragraph 3");

        ASSERT_EQ(u" LISTNUM  \\l 3", field->GetFieldCode());

        field = System::DynamicCast<FieldListNum>(builder->InsertField(FieldType::FieldListNum, true));

        // Setting this property to this particular value will emulate the AUTONUMOUT field
        field->set_ListName(u"OutlineDefault");

        ASSERT_TRUE(field->get_HasListName());
        ASSERT_EQ(u" LISTNUM  OutlineDefault", field->GetFieldCode());

        // Start counting from 1
        field->set_StartingNumber(u"1");
        builder->Writeln(u"Paragraph 4");

        // Our fields keep track of the count automatically, but the ListName needs to be set with each new field
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

        TestUtil::VerifyField(FieldType::FieldListNum, u" LISTNUM  \\l 3", String::Empty, field);
        ASSERT_TRUE(field->get_StartingNumber() == nullptr);
        ASSERT_EQ(u"3", field->get_ListLevel());
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
    //ExSummary:Shows how to insert a TOC and populate it with entries based on heading styles.
    void FieldToc_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // The table of contents we will insert will accept entries that are only within the scope of this bookmark
        builder->StartBookmark(u"MyBookmark");

        // Insert a list num field using a document builder
        auto field = System::DynamicCast<FieldToc>(builder->InsertField(FieldType::FieldTOC, true));

        // Limit possible TOC entries to only those within the bookmark we name here
        field->set_BookmarkName(u"MyBookmark");

        // Normally paragraphs with a "Heading n" style will be the only ones that will be added to a TOC as entries
        // We can set this attribute to include other styles, such as "Quote" and "Intense Quote" in this case
        field->set_CustomStyles(u"Quote; 6; Intense Quote; 7");

        // Styles are normally separated by a comma (",") but we can use this property to set a custom delimiter
        doc->get_FieldOptions()->set_CustomTocStyleSeparator(u";");

        // Filter out any headings that are outside this range
        field->set_HeadingLevelRange(u"1-3");

        // Headings in this range will not display their page number in their TOC entry
        field->set_PageNumberOmittingLevelRange(u"2-5");

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

        // These two headings will have the page numbers omitted because they are within the "2-5" range
        InsertNewPageWithHeading(builder, u"Fifth entry", u"Heading 2");
        InsertNewPageWithHeading(builder, u"Sixth entry", u"Heading 3");

        // This entry will be omitted because "Heading 4" is outside of the "1-3" range we set earlier
        InsertNewPageWithHeading(builder, u"Seventh entry", u"Heading 4");

        builder->EndBookmark(u"MyBookmark");
        builder->Writeln(u"Paragraph text.");

        // This entry will be omitted because it is outside the bookmark specified by the TOC
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

protected:
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

public:
    //ExStart
    //ExFor:FieldToc.EntryIdentifier
    //ExFor:FieldToc.EntryLevelRange
    //ExFor:FieldTC
    //ExFor:FieldTC.OmitPageNumber
    //ExFor:FieldTC.Text
    //ExFor:FieldTC.TypeIdentifier
    //ExFor:FieldTC.EntryLevel
    //ExSummary:Shows how to insert a TOC field and filter which TC fields end up as entries.
    void FieldTocEntryIdentifier()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartBookmark(u"MyBookmark");

        // Insert a list num field using a document builder
        auto fieldToc = System::DynamicCast<FieldToc>(builder->InsertField(FieldType::FieldTOC, true));
        fieldToc->set_EntryIdentifier(u"A");
        fieldToc->set_EntryLevelRange(u"1-3");

        ASSERT_EQ(u" TOC  \\f A \\l 1-3", fieldToc->GetFieldCode());

        // These two entries will appear in the table
        builder->InsertBreak(BreakType::PageBreak);
        InsertTocEntry(builder, u"TC field 1", u"A", u"1");
        InsertTocEntry(builder, u"TC field 2", u"A", u"2");

        ASSERT_EQ(u" TC  \"TC field 1\" \\n \\f A \\l 1", doc->get_Range()->get_Fields()->idx_get(1)->GetFieldCode());

        // These two entries will be omitted because of an incorrect type identifier
        InsertTocEntry(builder, u"TC field 3", u"B", u"1");

        // ...and an out-of-range entry level
        InsertTocEntry(builder, u"TC field 4", u"A", u"5");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.TC.docx");
        TestFieldTocEntryIdentifier(doc);
        //ExSkip
    }

    /// <summary>
    /// Insert a table of contents entry via a document builder.
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

protected:
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

public:
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

        // Insert a TOC field that creates a table of contents entry for each paragraph
        // that contains a SEQ field with a sequence identifier of "MySequence" with the number of the page which contains that field
        auto fieldToc = System::DynamicCast<FieldToc>(builder->InsertField(FieldType::FieldTOC, true));
        fieldToc->set_TableOfFiguresLabel(u"MySequence");

        // This identifier is for a parallel SEQ sequence,
        // the number that it is at will be displayed in front of the page number of the paragraph with the other sequence,
        // separated by a sequence separator character also defined below
        fieldToc->set_PrefixedSequenceIdentifier(u"PrefixSequence");
        fieldToc->set_SequenceSeparator(u">");

        ASSERT_EQ(u" TOC  \\c MySequence \\s PrefixSequence \\d >", fieldToc->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);

        // Insert a SEQ field to increment the sequence counter of "PrefixSequence" to 1
        // Since this paragraph doesn't contain a SEQ field of the "MySequence" sequence,
        // this will not appear as an entry in the TOC
        auto fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"PrefixSequence");
        builder->InsertParagraph();

        ASSERT_EQ(u" SEQ  PrefixSequence", fieldSeq->GetFieldCode());

        // Insert two SEQ fields, one for each of the sequences we defined above
        // The "MySequence" SEQ appears on page 2 and the "PrefixSequence" is at number 1 in this paragraph,
        // which means that our TOC will display this as an entry with the contents on the left and "1>2" on the right
        builder->Write(u"First TOC entry, MySequence #");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.TOC.SEQ.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.TOC.SEQ.docx");

        ASSERT_EQ(5, doc->get_Range()->get_Fields()->get_Count());

        fieldToc = System::DynamicCast<FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldTOC, u" TOC  \\c MySequence \\s PrefixSequence \\d >",
                              u"First TOC entry, MySequence #1\t\u0013 SEQ PrefixSequence _Toc256000000 \\* ARABIC \u00141\u0015>\u0013 PAGEREF _Toc256000000 "
                              u"\\h \u00142\u0015\r",
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

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  PrefixSequence", u"1", fieldSeq);
        ASSERT_EQ(u"PrefixSequence", fieldSeq->get_SequenceIdentifier());

        fieldSeq = System::DynamicCast<FieldSeq>(doc->get_Range()->get_Fields()->idx_get(4));

        TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"1", fieldSeq);
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
        //ExSummary:Shows how to reset numbering of a SEQ field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set the current number of the sequence to 100
        builder->Write(u"#");
        auto fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        fieldSeq->set_ResetNumber(u"100");

        ASSERT_EQ(u" SEQ  MySequence \\r 100", fieldSeq->GetFieldCode());

        builder->Write(u", #");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");

        // Insert a heading
        builder->InsertBreak(BreakType::ParagraphBreak);
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
        builder->Writeln(u"This level 1 heading will reset MySequence to 1");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Normal"));

        // Reset the sequence back to 1 when we encounter a heading of a specified level, which in this case is "1", same as the heading above
        builder->Write(u"\n#");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        fieldSeq->set_ResetHeadingLevel(u"1");

        ASSERT_EQ(u" SEQ  MySequence \\s 1", fieldSeq->GetFieldCode());

        // Move to the next number
        builder->Write(u", #");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        fieldSeq->set_InsertNextNumber(true);

        ASSERT_EQ(u" SEQ  MySequence \\n", fieldSeq->GetFieldCode());

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

        // This TOC takes in all SEQ fields with "MySequence" inside "TOCBookmark"
        auto fieldToc = System::DynamicCast<FieldToc>(builder->InsertField(FieldType::FieldTOC, true));
        fieldToc->set_TableOfFiguresLabel(u"MySequence");
        fieldToc->set_BookmarkName(u"TOCBookmark");
        builder->InsertBreak(BreakType::PageBreak);

        ASSERT_EQ(u" TOC  \\c MySequence \\b TOCBookmark", fieldToc->GetFieldCode());

        builder->Write(u"MySequence #");
        auto fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        builder->Writeln(u", won't show up in the TOC because it is outside of the bookmark.");

        builder->StartBookmark(u"TOCBookmark");

        builder->Write(u"MySequence #");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        builder->Writeln(u", will show up in the TOC next to the entry for the above caption.");

        builder->Write(u"MySequence #");
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"OtherSequence");
        builder->Writeln(u", won't show up in the TOC because it's from a different sequence identifier.");

        // The contents of the bookmark we reference here will not appear at the SEQ field, but will appear in the corresponding TOC entry
        fieldSeq = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        fieldSeq->set_SequenceIdentifier(u"MySequence");
        fieldSeq->set_BookmarkName(u"SEQBookmark");
        ASSERT_EQ(u" SEQ  MySequence SEQBookmark", fieldSeq->GetFieldCode());

        // Add bookmark to reference
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

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(String s)> _local_func_8 = [](String s)
        {
            return s.StartsWith(u"_Toc");
        };

        ArrayPtr<String> pageRefIds = fieldToc->get_Result()
                                                          .Split(MakeArray<char16_t>({u' '}))
                                                          ->LINQ_Where(static_cast<System::Func<String, bool>>(_local_func_8))
                                                          ->LINQ_ToArray();

        ASSERT_EQ(FieldType::FieldTOC, fieldToc->get_Type());
        ASSERT_EQ(u"MySequence", fieldToc->get_TableOfFiguresLabel());
        TestUtil::VerifyField(
            FieldType::FieldTOC, u" TOC  \\c MySequence \\b TOCBookmark",
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
        // Open a document that has bibliographical sources
        auto doc = MakeObject<Document>(MyDir + u"Bibliography.docx");
        ASSERT_EQ(2, doc->get_Range()->get_Fields()->get_Count());
        //ExSkip

        // Add text that we can cite
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Text to be cited with one source.");

        // Create a citation field using the document builder
        auto fieldCitation = System::DynamicCast<FieldCitation>(builder->InsertField(FieldType::FieldCitation, true));

        // A simple citation can have just the page number and author's name
        fieldCitation->set_SourceTag(u"Book1");
        // We refer to sources using their tag names
        fieldCitation->set_PageNumber(u"85");
        fieldCitation->set_SuppressAuthor(false);
        fieldCitation->set_SuppressTitle(true);
        fieldCitation->set_SuppressYear(true);

        ASSERT_EQ(u" CITATION  Book1 \\p 85 \\t \\y", fieldCitation->GetFieldCode());

        // We can make a more detailed citation and make it cite 2 sources
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

        // Insert a new page which will contain our bibliography
        builder->InsertBreak(BreakType::PageBreak);

        // All our sources can be displayed using a BIBLIOGRAPHY field
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
        //ExSummary:Shows how to insert a data field into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to insert a data field
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
        //ExSummary:Shows how to create an INCLUDE field and set its properties.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add an INCLUDE field with document builder and import a portion of the document defined by a bookmark
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

        auto fieldIncludePicture = System::DynamicCast<FieldIncludePicture>(builder->InsertField(FieldType::FieldIncludePicture, true));
        fieldIncludePicture->set_SourceFullName(ImageDir + u"Transparent background logo.png");

        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(fieldIncludePicture->GetFieldCode(), u" INCLUDEPICTURE  .*")->get_Success());

        // Here we apply the PNG32.FLT filter
        fieldIncludePicture->set_GraphicFilter(u"PNG32");
        fieldIncludePicture->set_IsLinked(true);
        fieldIncludePicture->set_ResizeHorizontally(true);
        fieldIncludePicture->set_ResizeVertically(true);

        // We can do the same thing with an IMPORT field
        auto fieldImport = System::DynamicCast<FieldImport>(builder->InsertField(FieldType::FieldImport, true));
        fieldImport->set_SourceFullName(ImageDir + u"Transparent background logo.png");
        fieldImport->set_GraphicFilter(u"PNG32");
        fieldImport->set_IsLinked(true);

        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(fieldImport->GetFieldCode(), u" IMPORT  .* \\\\c PNG32 \\\\d")->get_Success());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.INCLUDEPICTURE.docx");
        //ExEnd

        ASSERT_EQ(ImageDir + u"Transparent background logo.png", fieldIncludePicture->get_SourceFullName());
        ASSERT_EQ(u"PNG32", fieldIncludePicture->get_GraphicFilter());
        ASSERT_TRUE(fieldIncludePicture->get_IsLinked());
        ASSERT_TRUE(fieldIncludePicture->get_ResizeHorizontally());
        ASSERT_TRUE(fieldIncludePicture->get_ResizeVertically());

        ASSERT_EQ(ImageDir + u"Transparent background logo.png", fieldImport->get_SourceFullName());
        ASSERT_EQ(u"PNG32", fieldImport->get_GraphicFilter());
        ASSERT_TRUE(fieldImport->get_IsLinked());

        doc = MakeObject<Document>(ArtifactsDir + u"Field.INCLUDEPICTURE.docx");

        // The INCLUDEPICTURE fields have been converted into shapes with linked images during loading
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
    //ExSummary:Shows how to create an INCLUDETEXT field and set its properties.
    void FieldIncludeText_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to insert an include text field and perform an XSL transformation on an XML document
        SharedPtr<FieldIncludeText> fieldIncludeText =
            CreateFieldIncludeText(builder, MyDir + u"CD collection data.xml", false, u"text/xml", u"XML", u"ISO-8859-1");
        fieldIncludeText->set_XslTransformation(MyDir + u"CD collection XSL transformation.xsl");

        builder->Writeln();

        // Use a document builder to insert an include text field and use an XPath to take specific elements
        fieldIncludeText = CreateFieldIncludeText(builder, MyDir + u"CD collection data.xml", false, u"text/xml", u"XML", u"ISO-8859-1");
        fieldIncludeText->set_NamespaceMappings(u"xmlns:n='myNamespace'");
        fieldIncludeText->set_XPath(u"/catalog/cd/title");

        doc->Save(ArtifactsDir + u"Field.INCLUDETEXT.docx");
        TestFieldIncludeText(MakeObject<Document>(ArtifactsDir + u"Field.INCLUDETEXT.docx"));
        //ExSkip
    }

    /// <summary>
    /// Use a document builder to insert an INCLUDETEXT field and set its properties.
    /// </summary>
    SharedPtr<FieldIncludeText> CreateFieldIncludeText(SharedPtr<DocumentBuilder> builder, String sourceFullName, bool lockFields,
                                                               String mimeType, String textConverter, String encoding)
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

protected:
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

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        auto manager = MakeObject<System::Xml::XmlNamespaceManager>(cdCollectionXslTransformation->get_NameTable());
        manager->AddNamespace(u"xsl", u"http://www.w3.org/1999/XSL/Transform");

        for (int i = 0; i < table->get_Rows()->get_Count(); i++)
        {
            for (int j = 0; j < table->get_Rows()->idx_get(i)->get_Count(); j++)
            {
                if (i == 0)
                {
                    // When on the first row from the input document's table, ensure that all table's cells match all XML element Names
                    for (int k = 0; k < table->get_Rows()->get_Count() - 1; k++)
                    {
                        ASSERT_EQ(
                            catalogData->get_ChildNodes()->idx_get(k)->get_ChildNodes()->idx_get(j)->get_Name(),
                            table->get_Rows()->idx_get(i)->get_Cells()->idx_get(j)->GetText().Replace(ControlChar::Cell(), String::Empty).ToLower());
                    }

                    // Also make sure that the whole first row has the same color as the XSL transform
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
                    // When on all other rows of the input document's table, ensure that cell contents match XML element Values
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

public:
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
        //ExSummary:Shows how to insert HYPERLINK fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a hyperlink with a document builder
        auto field = System::DynamicCast<FieldHyperlink>(builder->InsertField(FieldType::FieldHyperlink, true));

        // When link is clicked, open a document and place the cursor on the bookmarked location
        field->set_Address(MyDir + u"Bookmarks.docx");
        field->set_SubAddress(u"MyBookmark3");
        field->set_ScreenTip(String(u"Open ") + field->get_Address() + u" on bookmark " + field->get_SubAddress() + u" in a new window");

        builder->Writeln();

        // Open html file at a specific frame
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
        ASSERT_EQ(String(u"Open ") + field->get_Address().Replace(u"\\", String::Empty) + u" on bookmark " + field->get_SubAddress() +
                      u" in a new window",
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

private:
    /// <summary>
    /// Image merging callback that pairs image shorthand names with filenames.
    /// </summary>
    class ImageFilenameCallback : public IFieldMergingCallback
    {
    public:
        ImageFilenameCallback()
        {
            // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            System::Details::ThisProtector __local_self_ref(this);
            //---------------------------------------------------------Self Reference

            mImageFilenames = MakeObject<System::Collections::Generic::Dictionary<String, String>>();
            mImageFilenames->Add(u"Dark logo", ImageDir + u"Logo.jpg");
            mImageFilenames->Add(u"Transparent logo", ImageDir + u"Transparent background logo.png");
        }

        void FieldMerging(SharedPtr<FieldMergingArgs> args) override
        {
            throw System::NotImplementedException();
        }

        void ImageFieldMerging(SharedPtr<ImageFieldMergingArgs> args) override
        {
            if (mImageFilenames->ContainsKey(System::ObjectExt::ToString(args->get_FieldValue())))
            {
                args->set_Image(System::Drawing::Image::FromFile(mImageFilenames->idx_get(System::ObjectExt::ToString(args->get_FieldValue()))));
            }

            ASSERT_FALSE(args->get_Image() == nullptr);
        }

    private:
        SharedPtr<System::Collections::Generic::Dictionary<String, String>> mImageFilenames;
    };
    //ExEnd

protected:
    void TestMergeFieldImages(SharedPtr<Document> doc)
    {
        doc = DocumentHelper::SaveOpen(doc);

        ASSERT_EQ(0, doc->get_Range()->get_Fields()->get_Count());
        ASSERT_EQ(2, doc->GetChildNodes(NodeType::Shape, true)->get_Count());

        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, shape);
        ASPOSE_ASSERT_EQ(300.0, shape->get_Width());
        ASPOSE_ASSERT_EQ(300.0, shape->get_Height());

        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 1, true));

        TestUtil::VerifyImageInShape(400, 400, ImageType::Png, shape);
        ASSERT_NEAR(300.0, shape->get_Width(), 1);
        ASSERT_NEAR(300.0, shape->get_Height(), 1);
    }

public:
    void FieldIndexFilter()
    {
        //ExStart
        //ExFor:FieldIndex
        //ExFor:FieldIndex.BookmarkName
        //ExFor:FieldIndex.EntryType
        //ExFor:FieldXE
        //ExFor:FieldXE.EntryType
        //ExFor:FieldXE.Text
        //ExSummary:Shows how to omit entries while populating an INDEX field with entries from XE fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        // Set these attributes so that an XE field shows up in the INDEX field's result
        // only if it is within the bounds of a bookmark named "MainBookmark", and is of type "A"
        index->set_BookmarkName(u"MainBookmark");
        index->set_EntryType(u"A");

        ASSERT_EQ(u" INDEX  \\b MainBookmark \\f A", index->GetFieldCode());

        // Our index will take up the first page
        builder->InsertBreak(BreakType::PageBreak);

        // Start the bookmark that will contain all eligible XE entries
        builder->StartBookmark(u"MainBookmark");

        // This entry will be picked up by the INDEX field because it is inside the bookmark
        // and its type matches the INDEX field's type
        // Note that even though the type is a string, it is defined by only the first character
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Index entry 1");
        indexEntry->set_EntryType(u"A");

        ASSERT_EQ(u" XE  \"Index entry 1\" \\f A", indexEntry->GetFieldCode());

        // Insert an XE field that will not appear in the INDEX field because it is of the wrong type
        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Index entry 2");
        indexEntry->set_EntryType(u"B");

        // End the bookmark and insert an XE field afterwards
        // It is of the same type as the INDEX field, but will not appear since it is outside of the bookmark
        // Note that the INDEX field itself does not have to be within its bookmark
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
        //ExSummary:Shows how to modify an INDEX field's appearance while populating it with XE field entries.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));
        index->set_LanguageId(u"1033");

        // Setting this attribute's value to "A" will group all the entries by their first letter
        // and place that letter in uppercase above each group
        index->set_Heading(u"A");

        // Set the table created by the INDEX field to span over 2 columns
        index->set_NumberOfColumns(u"2");

        // Set any entries with starting letters outside the "a-c" character range to be omitted
        index->set_LetterRange(u"a-c");

        ASSERT_EQ(u" INDEX  \\z 1033 \\h A \\c 2 \\p a-c", index->GetFieldCode());

        // These next two XE fields will show up under the "A" heading,
        // with their respective text stylings also applied to their page numbers
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

        // Both the next two XE fields will be under a "B" and "C" heading in the INDEX fields table of contents
        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Banana");

        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Cherry");

        // All INDEX field entries are sorted alphabetically, so this entry will show up under "A" with the other two
        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Avocado");

        // This entry will be excluded because, starting with the letter "D", it is outside the "a-c" range
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
        //ExSummary:Shows how to split a document into sections by combining INDEX and SEQ fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        // Set these two attributes to get the INDEX field's table of contents
        // to place the number that the "MySeq" sequence is at in each XE entry's location before the entry's page number,
        // separated by a custom character
        // Note that PageNumberSeparator and SequenceSeparator cannot be longer than 15 characters
        index->set_SequenceName(u"MySequence");
        index->set_PageNumberSeparator(u"\tMySequence at ");
        index->set_SequenceSeparator(u" on page ");
        ASSERT_TRUE(index->get_HasSequenceName());

        ASSERT_EQ(u" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index->GetFieldCode());

        // Insert a SEQ field which moves the "MySequence" sequence to 1
        // This field is treated as normal document text and will not show up on an INDEX field's table of contents
        builder->InsertBreak(BreakType::PageBreak);
        auto sequenceField = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        sequenceField->set_SequenceIdentifier(u"MySequence");

        ASSERT_EQ(u" SEQ  MySequence", sequenceField->GetFieldCode());

        // Insert a XE field which will show up in the INDEX field
        // Since "MySequence" is at 1 and this XE field is on page 2, along with the custom separators we defined above,
        // this field's INDEX entry will say "MySequence at 1 on page 2"
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Cat");

        ASSERT_EQ(u" XE  Cat", indexEntry->GetFieldCode());

        // Insert a page break and advance "MySequence" by 2
        builder->InsertBreak(BreakType::PageBreak);
        sequenceField = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        sequenceField->set_SequenceIdentifier(u"MySequence");
        sequenceField = System::DynamicCast<FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
        sequenceField->set_SequenceIdentifier(u"MySequence");

        // Insert a XE field with the same text as the one above, which will thus be appended to the same entry in the INDEX field
        // Since we are on page 2 with "MySequence" at 3, ", 3 on page 3" will be appended to the same INDEX entry as above
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Cat");

        // Insert an XE field which makes a new entry with MySequence at 3 on page 4
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

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Aspose::Words::Fields::Field> f)> _local_func_9 = [](SharedPtr<Aspose::Words::Fields::Field> f)
        {
            return f->get_Type() == FieldType::FieldSequence;
        };

        ASSERT_EQ(3, doc->get_Range()->get_Fields()->LINQ_Where(static_cast<System::Func<SharedPtr<Field>, bool>>(_local_func_9))->LINQ_Count());
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

        // Create an INDEX field which will display a table with the page locations of XE fields in the document body
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        // Set a page number separator and a page number separator
        // The page number separator will go between the INDEX entry's name and first page a corresponsing XE field appears,
        // while the page number list separator will appear between page numbers if there are multiple in the same INDEX field entry
        index->set_PageNumberSeparator(u", on page(s) ");
        index->set_PageNumberListSeparator(u" & ");

        ASSERT_EQ(u" INDEX  \\e \", on page(s) \" \\l \" & \"", index->GetFieldCode());
        ASSERT_TRUE(index->get_HasPageNumberSeparator());

        // Insert 3 XE entries with the same name on three different pages so they all end up in one INDEX field table entry,
        // where both our separators will be applied, resulting in a value of "First entry, on page(s) 2 & 3 & 4"
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
        //ExFor:FieldXE.HasPageRangeBookmarkName
        //ExFor:FieldXE.PageRangeBookmarkName
        //ExSummary:Shows how to specify a bookmark's spanned pages as a page range for an INDEX field entry.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        index->set_PageNumberSeparator(u", on page(s) ");
        index->set_PageRangeSeparator(u" to ");

        ASSERT_EQ(u" INDEX  \\e \", on page(s) \" \\g \" to \"", index->GetFieldCode());

        // Insert an XE field on page 2
        builder->InsertBreak(BreakType::PageBreak);
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"My entry");

        // If we use this attribute to refer to a bookmark,
        // this XE field's page number will be substituted by the page range that the referenced bookmark spans
        indexEntry->set_PageRangeBookmarkName(u"MyBookmark");

        ASSERT_EQ(u" XE  \"My entry\" \\r MyBookmark", indexEntry->GetFieldCode());
        ASSERT_TRUE(indexEntry->get_HasPageRangeBookmarkName());

        // Insert a bookmark that starts on page 3 and ends on page 5
        // Since the XE field references this bookmark,
        // its location page number will show up in the INDEX field's table as "3 to 5" instead of "2",
        // which is its actual page
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
        ASSERT_TRUE(indexEntry->get_HasPageRangeBookmarkName());
    }

    void FieldIndexCrossReferenceSeparator()
    {
        //ExStart
        //ExFor:FieldIndex.CrossReferenceSeparator
        //ExFor:FieldXE.PageNumberReplacement
        //ExSummary:Shows how to define cross references in an INDEX field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        // Define a custom separator that is applied if an XE field contains a page number replacement
        index->set_CrossReferenceSeparator(u", see: ");

        ASSERT_EQ(u" INDEX  \\k \", see: \"", index->GetFieldCode());

        // Insert an XE field on page 2
        // That page number, together with the field's Text attribute, will show up as a table of contents entry in the INDEX field,
        builder->InsertBreak(BreakType::PageBreak);
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Apple");

        ASSERT_EQ(u" XE  Apple", indexEntry->GetFieldCode());

        // Insert another XE field on page 3, and set a value for "PageNumberReplacement"
        // In the INDEX field's table, this field will display the value of that attribute after the field's CrossReferenceSeparator instead of the page number
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

    void FieldIndexSubheading(bool doRunSubentriesOnTheSameLine)
    {
        //ExStart
        //ExFor:FieldIndex.RunSubentriesOnSameLine
        //ExSummary:Shows how to work with subentries in an INDEX field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        // Normally, every XE field that's a subheading of any level is displayed on a unique line entry
        // in the INDEX field's table of contents
        // We can reduce the length of our INDEX table by putting all subheading entries along with their page locations on one line
        index->set_RunSubentriesOnSameLine(doRunSubentriesOnTheSameLine);
        index->set_PageNumberSeparator(u", see page ");
        index->set_Heading(u"A");

        if (doRunSubentriesOnTheSameLine)
        {
            ASSERT_EQ(u" INDEX  \\r \\e \", see page \" \\h A", index->GetFieldCode());
        }
        else
        {
            ASSERT_EQ(u" INDEX  \\e \", see page \" \\h A", index->GetFieldCode());
        }

        // An XE field's "Text" attribute is the same thing as the "Heading" that will appear in the INDEX field's table of contents
        // This attribute can also contain one or multiple subheadings, separated by a colon (:),
        // which will be grouped under their parent headings/subheadings in the INDEX field
        // If index.RunSubentriesOnSameLine is false, "Heading 1" will take up one line as a heading,
        // followed by a two-line indented list of "Subheading 1" and "Subheading 2" with their respective page numbers
        // Otherwise, the two subheadings and their page numbers will be on the same line as their heading
        builder->InsertBreak(BreakType::PageBreak);
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Heading 1:Subheading 1");

        ASSERT_EQ(u" XE  \"Heading 1:Subheading 1\"", indexEntry->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);
        indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"Heading 1:Subheading 2");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.INDEX.XE.Subheading.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.INDEX.XE.Subheading.docx");
        index = System::DynamicCast<FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));

        if (doRunSubentriesOnTheSameLine)
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

    void FieldIndexYomi(bool doSortEntriesUsingYomi)
    {
        //ExStart
        //ExFor:FieldIndex.UseYomi
        //ExFor:FieldXE.Yomi
        //ExSummary:Shows how to sort INDEX field entries phonetically.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        auto index = System::DynamicCast<FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));

        // Set the INDEX table to sort entries phonetically using Hiragana
        index->set_UseYomi(doSortEntriesUsingYomi);

        if (doSortEntriesUsingYomi)
        {
            ASSERT_EQ(u" INDEX  \\y", index->GetFieldCode());
        }
        else
        {
            ASSERT_EQ(u" INDEX ", index->GetFieldCode());
        }

        // Insert 4 XE fields, which would show up as entries in the INDEX field's table of contents,
        // sorted in lexicographic order on their "Text" attribute
        builder->InsertBreak(BreakType::PageBreak);
        auto indexEntry = System::DynamicCast<FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
        indexEntry->set_Text(u"愛子");

        // The "Text" attribute may contain a word's spelling in Kanji, whose pronunciation may be ambiguous,
        // while a "Yomi" version of the word will be spelled exactly how it is pronounced using Hiragana
        // If our INDEX field is set to use Yomi, then we can sort phonetically using the "Yomi" attribute values instead of the "Text" attribute
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

        if (doSortEntriesUsingYomi)
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
        //ExSummary:Shows how to insert a BARCODE field and set its properties.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to insert a bookmark with a US postal code in it
        builder->StartBookmark(u"BarcodeBookmark");
        builder->Writeln(u"96801");
        builder->EndBookmark(u"BarcodeBookmark");

        builder->Writeln();

        // Reference a US postal code directly
        auto field = System::DynamicCast<FieldBarcode>(builder->InsertField(FieldType::FieldBarcode, true));
        field->set_FacingIdentificationMark(u"C");
        field->set_PostalAddress(u"96801");
        field->set_IsUSPostalAddress(true);

        ASSERT_EQ(u" BARCODE  96801 \\f C \\u", field->GetFieldCode());

        builder->Writeln();

        // Reference a US postal code from a bookmark
        field = System::DynamicCast<FieldBarcode>(builder->InsertField(FieldType::FieldBarcode, true));
        field->set_PostalAddress(u"BarcodeBookmark");
        field->set_IsBookmark(true);

        ASSERT_EQ(u" BARCODE  BarcodeBookmark \\b", field->GetFieldCode());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.BARCODE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.BARCODE.docx");

        ASSERT_EQ(0, doc->GetChildNodes(NodeType::Shape, true)->get_Count());

        field = System::DynamicCast<FieldBarcode>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldBarcode, u" BARCODE  96801 \\f C \\u", String::Empty, field);
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
        //ExSummary:Shows how to insert a DISPLAYBARCODE field and set its properties.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto field = System::DynamicCast<FieldDisplayBarcode>(builder->InsertField(FieldType::FieldDisplayBarcode, true));

        // Insert a QR code
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

        // Insert an EAN13 barcode
        field = System::DynamicCast<FieldDisplayBarcode>(builder->InsertField(FieldType::FieldDisplayBarcode, true));
        field->set_BarcodeType(u"EAN13");
        field->set_BarcodeValue(u"501234567890");
        field->set_DisplayText(true);
        field->set_PosCodeStyle(u"CASE");
        field->set_FixCheckDigit(true);

        ASSERT_EQ(u" DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", field->GetFieldCode());
        builder->Writeln();

        // Insert a CODE39 barcode
        field = System::DynamicCast<FieldDisplayBarcode>(builder->InsertField(FieldType::FieldDisplayBarcode, true));
        field->set_BarcodeType(u"CODE39");
        field->set_BarcodeValue(u"12345ABCDE");
        field->set_AddStartStopChar(true);

        ASSERT_EQ(u" DISPLAYBARCODE  12345ABCDE CODE39 \\d", field->GetFieldCode());
        builder->Writeln();

        // Insert an ITF14 barcode
        field = System::DynamicCast<FieldDisplayBarcode>(builder->InsertField(FieldType::FieldDisplayBarcode, true));
        field->set_BarcodeType(u"ITF14");
        field->set_BarcodeValue(u"09312345678907");
        field->set_CaseCodeStyle(u"STD");

        ASSERT_EQ(u" DISPLAYBARCODE  09312345678907 ITF14 \\c STD", field->GetFieldCode());

        doc->UpdateFields();
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
    //ExSummary:Shows how to insert linked objects as LINK, DDE and DDEAUTO fields and present them within the document in different ways.
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

        // Insert fields containing text from another document and present them as text (see InsertLinkedObjectAs enum)
        builder->Writeln(u"FieldLink:\n");
        InsertFieldLink(builder, insertLinkedObjectAs, u"Word.Document.8", MyDir + u"Document.docx", nullptr, true);

        builder->Writeln(u"FieldDde:\n");
        InsertFieldDde(builder, insertLinkedObjectAs, u"Excel.Sheet", MyDir + u"Spreadsheet.xlsx", u"Sheet1!R1C1", true, true);

        builder->Writeln(u"FieldDdeAuto:\n");
        InsertFieldDdeAuto(builder, insertLinkedObjectAs, u"Excel.Sheet", MyDir + u"Spreadsheet.xlsx", u"Sheet1!R1C1", true);

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.LINK.DDE.DDEAUTO.docx");
    }

    void FieldLinkedObjectsAsImage(ExField::InsertLinkedObjectAs insertLinkedObjectAs)
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert one cell from a spreadsheet as an image (see InsertLinkedObjectAs enum)
        builder->Writeln(u"FieldLink:\n");
        InsertFieldLink(builder, insertLinkedObjectAs, u"Excel.Sheet", MyDir + u"MySpreadsheet.xlsx", u"Sheet1!R2C2", true);

        builder->Writeln(u"FieldDde:\n");
        InsertFieldDde(builder, insertLinkedObjectAs, u"Excel.Sheet", MyDir + u"Spreadsheet.xlsx", u"Sheet1!R1C1", true, true);

        builder->Writeln(u"FieldDdeAuto:\n");
        InsertFieldDdeAuto(builder, insertLinkedObjectAs, u"Excel.Sheet", MyDir + u"Spreadsheet.xlsx", u"Sheet1!R1C1", true);

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.LINK.DDE.DDEAUTO.AsImage.docx");
    }

protected:
    /// <summary>
    /// Use a document builder to insert a LINK field and set its properties according to parameters.
    /// </summary>
    static void InsertFieldLink(SharedPtr<DocumentBuilder> builder, ExField::InsertLinkedObjectAs insertLinkedObjectAs, String progId,
                                String sourceFullName, String sourceItem, bool shouldAutoUpdate)
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
    /// Use a document builder to insert a DDE field and set its properties according to parameters.
    /// </summary>
    static void InsertFieldDde(SharedPtr<DocumentBuilder> builder, ExField::InsertLinkedObjectAs insertLinkedObjectAs, String progId,
                               String sourceFullName, String sourceItem, bool isLinked, bool shouldAutoUpdate)
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
    /// Use a document builder to insert a DDEAUTO field and set its properties according to parameters.
    /// </summary>
    static void InsertFieldDdeAuto(SharedPtr<DocumentBuilder> builder, ExField::InsertLinkedObjectAs insertLinkedObjectAs, String progId,
                                   String sourceFullName, String sourceItem, bool isLinked)
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

public:
    void FieldUserAddress_()
    {
        //ExStart
        //ExFor:FieldUserAddress
        //ExFor:FieldUserAddress.UserAddress
        //ExSummary:Shows how to use the USERADDRESS field.
        auto doc = MakeObject<Document>();

        // Create a user information object and set it as the data source for our field
        auto userInformation = MakeObject<UserInformation>();
        userInformation->set_Address(u"123 Main Street");
        doc->get_FieldOptions()->set_CurrentUser(userInformation);

        // Display the current user's address with a USERADDRESS field
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto fieldUserAddress = System::DynamicCast<FieldUserAddress>(builder->InsertField(FieldType::FieldUserAddress, true));
        ASSERT_EQ(userInformation->get_Address(), fieldUserAddress->get_Result());

        ASSERT_EQ(u" USERADDRESS ", fieldUserAddress->GetFieldCode());
        ASSERT_EQ(u"123 Main Street", fieldUserAddress->get_Result());

        // We can set this attribute to get our field to display a different value
        fieldUserAddress->set_UserAddress(u"456 North Road");
        fieldUserAddress->Update();

        ASSERT_EQ(u" USERADDRESS  \"456 North Road\"", fieldUserAddress->GetFieldCode());
        ASSERT_EQ(u"456 North Road", fieldUserAddress->get_Result());

        // This does not change the value in the user information object
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

        // Create a user information object and set it as the data source for our field
        auto userInformation = MakeObject<UserInformation>();
        userInformation->set_Initials(u"J. D.");
        doc->get_FieldOptions()->set_CurrentUser(userInformation);

        // Display the current user's Initials with a USERINITIALS field
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto fieldUserInitials = System::DynamicCast<FieldUserInitials>(builder->InsertField(FieldType::FieldUserInitials, true));
        ASSERT_EQ(userInformation->get_Initials(), fieldUserInitials->get_Result());

        ASSERT_EQ(u" USERINITIALS ", fieldUserInitials->GetFieldCode());
        ASSERT_EQ(u"J. D.", fieldUserInitials->get_Result());

        // We can set this attribute to get our field to display a different value
        fieldUserInitials->set_UserInitials(u"J. C.");
        fieldUserInitials->Update();

        ASSERT_EQ(u" USERINITIALS  \"J. C.\"", fieldUserInitials->GetFieldCode());
        ASSERT_EQ(u"J. C.", fieldUserInitials->get_Result());

        // This does not change the value in the user information object
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

        // Create a user information object and set it as the data source for our field
        auto userInformation = MakeObject<UserInformation>();
        userInformation->set_Name(u"John Doe");
        doc->get_FieldOptions()->set_CurrentUser(userInformation);

        auto builder = MakeObject<DocumentBuilder>(doc);

        // Display the current user's Name with a USERNAME field
        auto fieldUserName = System::DynamicCast<FieldUserName>(builder->InsertField(FieldType::FieldUserName, true));
        ASSERT_EQ(userInformation->get_Name(), fieldUserName->get_Result());

        ASSERT_EQ(u" USERNAME ", fieldUserName->GetFieldCode());
        ASSERT_EQ(u"John Doe", fieldUserName->get_Result());

        // We can set this attribute to get our field to display a different value
        fieldUserName->set_UserName(u"Jane Doe");
        fieldUserName->Update();

        ASSERT_EQ(u" USERNAME  \"Jane Doe\"", fieldUserName->GetFieldCode());
        ASSERT_EQ(u"Jane Doe", fieldUserName->get_Result());

        // This does not change the value in the user information object
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

        // Create a list based on one of the Microsoft Word list templates
        SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(Lists::ListTemplate::NumberDefault);

        // This generated list will look like "1.a )"
        // The space before the bracket is a non-delimiter character that can be suppressed
        list->get_ListLevels()->idx_get(0)->set_NumberFormat(u"\x0000"
                                                             u".");
        list->get_ListLevels()->idx_get(1)->set_NumberFormat(u"\x0001"
                                                             u" )");

        // Add text and apply paragraph styles that will be referenced by STYLEREF fields
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

        // Place a STYLEREF field in the header and have it display the first "List Paragraph"-styled text in the document
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        auto field = System::DynamicCast<FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
        field->set_StyleName(u"List Paragraph");

        // Place a STYLEREF field in the footer and have it display the last text
        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        field = System::DynamicCast<FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
        field->set_StyleName(u"List Paragraph");
        field->set_SearchFromBottom(true);

        builder->MoveToDocumentEnd();

        // We can also use STYLEREF fields to reference the list numbers of lists
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
        //ExSummary:Shows how to insert DATE fields with different kinds of calendars.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // One way of putting dates into our documents is inserting DATE fields with document builder
        auto field = System::DynamicCast<FieldDate>(builder->InsertField(FieldType::FieldDate, true));

        // Set the field's date to the current date of the Islamic Lunar Calendar
        field->set_UseLunarCalendar(true);
        ASSERT_EQ(u" DATE  \\h", field->GetFieldCode());
        builder->Writeln();

        // Insert a date field with the current date of the Umm al-Qura calendar
        field = System::DynamicCast<FieldDate>(builder->InsertField(FieldType::FieldDate, true));
        field->set_UseUmAlQuraCalendar(true);
        ASSERT_EQ(u" DATE  \\u", field->GetFieldCode());
        builder->Writeln();

        // Insert a date field with the current date of the Indian national calendar
        field = System::DynamicCast<FieldDate>(builder->InsertField(FieldType::FieldDate, true));
        field->set_UseSakaEraCalendar(true);
        ASSERT_EQ(u" DATE  \\s", field->GetFieldCode());
        builder->Writeln();

        // Insert a date field with the current date of the calendar used in the (Insert > Date and Time) dialog box
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
        //ExSummary:Shows how to insert CREATEDATE fields to display document creation dates.
        // Open an existing document and move a document builder to the end
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveToDocumentEnd();
        builder->Writeln(u" Date this document was created:");

        // Insert a CREATEDATE field and display, using the Lunar Calendar, the date the document was created
        builder->Write(u"According to the Lunar Calendar - ");
        auto field = System::DynamicCast<FieldCreateDate>(builder->InsertField(FieldType::FieldCreateDate, true));
        field->set_UseLunarCalendar(true);

        ASSERT_EQ(u" CREATEDATE  \\h", field->GetFieldCode());

        // Display the date using the Umm al-Qura Calendar
        builder->Write(u"\nAccording to the Umm al-Qura Calendar - ");
        field = System::DynamicCast<FieldCreateDate>(builder->InsertField(FieldType::FieldCreateDate, true));
        field->set_UseUmAlQuraCalendar(true);

        ASSERT_EQ(u" CREATEDATE  \\u", field->GetFieldCode());

        // Display the date using the Indian National Calendar
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
        //ExFor:FieldSaveDate
        //ExFor:FieldSaveDate.UseLunarCalendar
        //ExFor:FieldSaveDate.UseSakaEraCalendar
        //ExFor:FieldSaveDate.UseUmAlQuraCalendar
        //ExSummary:Shows how to insert SAVEDATE fields the date and time a document was last saved.
        // Open an existing document and move a document builder to the end
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveToDocumentEnd();
        builder->Writeln(u" Date this document was last saved:");

        // Insert a SAVEDATE field and display, using the Lunar Calendar, the date the document was last saved
        builder->Write(u"According to the Lunar Calendar - ");
        auto field = System::DynamicCast<FieldSaveDate>(builder->InsertField(FieldType::FieldSaveDate, true));
        field->set_UseLunarCalendar(true);

        ASSERT_EQ(u" SAVEDATE  \\h", field->GetFieldCode());

        // Display the date using the Umm al-Qura Calendar
        builder->Write(u"\nAccording to the Umm al-Qura calendar - ");
        field = System::DynamicCast<FieldSaveDate>(builder->InsertField(FieldType::FieldSaveDate, true));
        field->set_UseUmAlQuraCalendar(true);

        ASSERT_EQ(u" SAVEDATE  \\u", field->GetFieldCode());

        // Display the date using the Indian National Calendar
        builder->Write(u"\nAccording to the Indian National calendar - ");
        field = System::DynamicCast<FieldSaveDate>(builder->InsertField(FieldType::FieldSaveDate, true));
        field->set_UseSakaEraCalendar(true);

        ASSERT_EQ(u" SAVEDATE  \\s", field->GetFieldCode());

        // While the date/time of the most recent save operation is tracked automatically by Microsoft Word,
        // we will need to update the value manually if we wish to do the same thing when calling the Save() method
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
        //ExSummary:Shows how to insert fields using a field builder.
        auto doc = MakeObject<Document>();

        // Use a field builder to add a SYMBOL field which displays the "F with hook" symbol
        auto builder = MakeObject<FieldBuilder>(FieldType::FieldSymbol);
        builder->AddArgument(402);
        builder->AddSwitch(u"\\f", u"Arial");
        builder->AddSwitch(u"\\s", 25);
        builder->AddSwitch(u"\\u");
        SharedPtr<Field> field = builder->BuildAndInsert(doc->get_FirstSection()->get_Body()->get_FirstParagraph());

        ASSERT_EQ(u" SYMBOL 402 \\f Arial \\s 25 \\u ", field->GetFieldCode());

        // Use a field builder to create a formula field that will be used by another field builder
        auto innerFormulaBuilder = MakeObject<FieldBuilder>(FieldType::FieldFormula);
        innerFormulaBuilder->AddArgument(100);
        innerFormulaBuilder->AddArgument(u"+");
        innerFormulaBuilder->AddArgument(74);

        // Add a field builder as an argument to another field builder
        // The result of our formula field will be used as an ANSI value representing the "enclosed R" symbol,
        // to be displayed by this SYMBOL field
        builder = MakeObject<FieldBuilder>(FieldType::FieldSymbol);
        builder->AddArgument(innerFormulaBuilder);
        field = builder->BuildAndInsert(doc->get_FirstSection()->get_Body()->AppendParagraph(u""));

        ASSERT_EQ(u" SYMBOL \u0013 = 100 + 74 \u0014\u0015 ", field->GetFieldCode());

        // Now we will use our builder to construct a more complex field with nested fields
        // For our IF field, we will first create two formula fields to serve as expressions
        // Their results will be tested for equality to decide what value an IF field displays
        auto leftExpression = MakeObject<FieldBuilder>(FieldType::FieldFormula);
        leftExpression->AddArgument(2);
        leftExpression->AddArgument(u"+");
        leftExpression->AddArgument(3);

        auto rightExpression = MakeObject<FieldBuilder>(FieldType::FieldFormula);
        rightExpression->AddArgument(2.5);
        rightExpression->AddArgument(u"*");
        rightExpression->AddArgument(5.2);

        // Next, we will create two field arguments using field argument builders
        // These will serve as the two possible outputs of our IF field and they will also use our two expressions
        auto trueOutput = MakeObject<FieldArgumentBuilder>();
        trueOutput->AddText(u"True, both expressions amount to ");
        trueOutput->AddField(leftExpression);

        auto falseOutput = MakeObject<FieldArgumentBuilder>();
        falseOutput->AddNode(MakeObject<Run>(doc, u"False, "));
        falseOutput->AddField(leftExpression);
        falseOutput->AddNode(MakeObject<Run>(doc, u" does not equal "));
        falseOutput->AddField(rightExpression);

        // Finally, we will use a field builder to create an IF field which takes two field builders as expressions,
        // and two field argument builders as the two potential outputs
        builder = MakeObject<FieldBuilder>(FieldType::FieldIf);
        builder->AddArgument(leftExpression);
        builder->AddArgument(u"=");
        builder->AddArgument(rightExpression);
        builder->AddArgument(trueOutput);
        builder->AddArgument(falseOutput);

        builder->BuildAndInsert(doc->get_FirstSection()->get_Body()->AppendParagraph(u""));

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

        ASSERT_NE(
            doc->get_Range()->get_Fields()->idx_get(2)->get_Start()->get_ParentNode(),
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
        //ExSummary:Shows how to display a document creator's name with an AUTHOR field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // If we open an existing document, the document's author's full name will be displayed by the field
        // If we create a document programmatically, we need to set this attribute to the author's name, so our field has something to display
        doc->get_FieldOptions()->set_DefaultDocumentAuthor(u"Joe Bloggs");

        builder->Write(u"This document was created by ");
        auto field = System::DynamicCast<FieldAuthor>(builder->InsertField(FieldType::FieldAuthor, true));
        field->Update();

        ASSERT_EQ(u" AUTHOR ", field->GetFieldCode());
        ASSERT_EQ(u"Joe Bloggs", field->get_Result());

        // If this property has a value, it will supersede the one we set above
        doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
        field->Update();

        ASSERT_EQ(u" AUTHOR ", field->GetFieldCode());
        ASSERT_EQ(u"John Doe", field->get_Result());

        // Our field can also override the document's built in author name like this
        field->set_AuthorName(u"Jane Doe");
        field->Update();

        ASSERT_EQ(u" AUTHOR  \"Jane Doe\"", field->GetFieldCode());
        ASSERT_EQ(u"Jane Doe", field->get_Result());

        // The author name in the built-in properties was changed by the field, but the default document author stays the same
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
        //ExSummary:Shows how to use fields to display document properties and variables.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set the value of a document property
        doc->get_BuiltInDocumentProperties()->set_Category(u"My category");

        // Display the value of that property with a DOCPROPERTY field
        auto fieldDocProperty = System::DynamicCast<FieldDocProperty>(builder->InsertField(u" DOCPROPERTY Category "));
        fieldDocProperty->Update();

        ASSERT_EQ(u" DOCPROPERTY Category ", fieldDocProperty->GetFieldCode());
        ASSERT_EQ(u"My category", fieldDocProperty->get_Result());

        builder->Writeln();

        // While the set of a document's properties is fixed, we can add, name, and define our own values in the variables collection
        ASSERT_EQ(0, doc->get_Variables()->get_Count());
        doc->get_Variables()->Add(u"My variable", u"My variable's value");

        // We can access a variable using its name and display it with a DOCVARIABLE field
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

        // Set a value for the document's subject property
        doc->get_BuiltInDocumentProperties()->set_Subject(u"My subject");

        // We can display this value with a SUBJECT field
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto field = System::DynamicCast<FieldSubject>(builder->InsertField(FieldType::FieldSubject, true));
        field->Update();

        ASSERT_EQ(u" SUBJECT ", field->GetFieldCode());
        ASSERT_EQ(u"My subject", field->get_Result());

        // We can also set the field's Text attribute to override the current value of the Subject property
        field->set_Text(u"My new subject");
        field->Update();

        ASSERT_EQ(u" SUBJECT  \"My new subject\"", field->GetFieldCode());
        ASSERT_EQ(u"My new subject", field->get_Result());

        // As well as displaying a new value in our field, we also changed the value of the document property
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
        //ExSummary:Shows how to use the COMMENTS field to display a document's comments.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // This property is where the COMMENTS field will source its content from
        doc->get_BuiltInDocumentProperties()->set_Comments(u"My comment.");

        // Insert a COMMENTS field with a document builder
        auto field = System::DynamicCast<FieldComments>(builder->InsertField(FieldType::FieldComments, true));
        field->Update();

        ASSERT_EQ(u" COMMENTS ", field->GetFieldCode());
        ASSERT_EQ(u"My comment.", field->get_Result());

        // We can override the comment from the document's built in properties and display any text we put here instead
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
        // Open a document and verify its file size
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        ASSERT_EQ(16222, doc->get_BuiltInDocumentProperties()->get_Bytes());

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveToDocumentEnd();
        builder->InsertParagraph();

        // By default, file size is displayed in bytes
        auto field = System::DynamicCast<FieldFileSize>(builder->InsertField(FieldType::FieldFileSize, true));
        field->Update();

        ASSERT_EQ(u" FILESIZE ", field->GetFieldCode());
        ASSERT_EQ(u"16222", field->get_Result());

        // Set the field to display size in kilobytes
        builder->InsertParagraph();
        field = System::DynamicCast<FieldFileSize>(builder->InsertField(FieldType::FieldFileSize, true));
        field->set_IsInKilobytes(true);
        field->Update();

        ASSERT_EQ(u" FILESIZE  \\k", field->GetFieldCode());
        ASSERT_EQ(u"16", field->get_Result());

        // Set the field to display size in megabytes
        builder->InsertParagraph();
        field = System::DynamicCast<FieldFileSize>(builder->InsertField(FieldType::FieldFileSize, true));
        field->set_IsInMegabytes(true);
        field->Update();

        ASSERT_EQ(u" FILESIZE  \\m", field->GetFieldCode());
        ASSERT_EQ(u"0", field->get_Result());

        // To update the values of these fields while editing in Microsoft Word,
        // the changes must first be saved, then the fields need to be manually updated
        doc->Save(ArtifactsDir + u"Field.FILESIZE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.FILESIZE.docx");

        field = System::DynamicCast<FieldFileSize>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldFileSize, u" FILESIZE ", u"16222", field);

        // These fields will need to be updated to produce an accurate result
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

        // Add a GOTOBUTTON which will take us to a bookmark referenced by "MyBookmark"
        auto field = System::DynamicCast<FieldGoToButton>(builder->InsertField(FieldType::FieldGoToButton, true));
        field->set_DisplayText(u"My Button");
        field->set_Location(u"MyBookmark");

        ASSERT_EQ(u" GOTOBUTTON  MyBookmark My Button", field->GetFieldCode());

        // Add an arrival destination for our button
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

    void FieldFillIn_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a FILLIN field with a document builder
        auto field = System::DynamicCast<FieldFillIn>(builder->InsertField(FieldType::FieldFillIn, true));
        field->set_PromptText(u"Please enter a response:");
        field->set_DefaultResponse(u"A default response.");

        // Set this to prompt the user for a response when a mail merge is performed
        field->set_PromptOnceOnMailMerge(true);

        ASSERT_EQ(u" FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o", field->GetFieldCode());

        // Perform a simple mail merge
        auto mergeField = System::DynamicCast<FieldMergeField>(builder->InsertField(FieldType::FieldMergeField, true));
        mergeField->set_FieldName(u"MergeField");

        doc->get_FieldOptions()->set_UserPromptRespondent(MakeObject<ExField::PromptRespondent>());
        doc->get_MailMerge()->Execute(MakeArray<String>({u"MergeField"}),
                                      MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(u"")}));

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.FILLIN.docx");
        TestFieldFillIn(MakeObject<Document>(ArtifactsDir + u"Field.FILLIN.docx"));
        //ExSkip
    }

private:
    /// <summary>
    /// IFieldUserPromptRespondent implementation that appends a line to the default response of an FILLIN field during a mail merge.
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

protected:
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

public:
    void FieldInfo_()
    {
        //ExStart
        //ExFor:FieldInfo
        //ExFor:FieldInfo.InfoType
        //ExFor:FieldInfo.NewValue
        //ExSummary:Shows how to work with INFO fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set the value of a document property
        doc->get_BuiltInDocumentProperties()->set_Comments(u"My comment");

        // We can access a property using its name and display it with an INFO field
        // In this case, it will be the Comments property
        auto field = System::DynamicCast<FieldInfo>(builder->InsertField(FieldType::FieldInfo, true));
        field->set_InfoType(u"Comments");
        field->Update();

        ASSERT_EQ(u" INFO  Comments", field->GetFieldCode());
        ASSERT_EQ(u"My comment", field->get_Result());

        builder->Writeln();

        // We can override the value of a document property by setting an INFO field's optional new value
        field = System::DynamicCast<FieldInfo>(builder->InsertField(FieldType::FieldInfo, true));
        field->set_InfoType(u"Comments");
        field->set_NewValue(u"New comment");
        field->Update();

        // Our field's new value has been applied to the corresponding property
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
        //ExSummary:Shows how to use MACROBUTTON fields that enable us to run macros by clicking.
        // Open a document that contains macros
        auto doc = MakeObject<Document>(MyDir + u"Macro.docm");
        auto builder = MakeObject<DocumentBuilder>(doc);

        ASSERT_TRUE(doc->get_HasMacros());

        // Insert a MACROBUTTON field and reference by name a macro that exists within the input document
        auto field = System::DynamicCast<FieldMacroButton>(builder->InsertField(FieldType::FieldMacroButton, true));
        field->set_MacroName(u"MyMacro");
        field->set_DisplayText(String(u"Double click to run macro: ") + field->get_MacroName());

        ASSERT_EQ(u" MACROBUTTON  MyMacro Double click to run macro: MyMacro", field->GetFieldCode());

        // Reference "ViewZoom200", a macro that was shipped with Microsoft Word, found under "Word commands"
        // If our document has a macro of the same name as one from another source, the field will select ours to run
        builder->InsertParagraph();
        field = System::DynamicCast<FieldMacroButton>(builder->InsertField(FieldType::FieldMacroButton, true));
        field->set_MacroName(u"ViewZoom200");
        field->set_DisplayText(String(u"Run ") + field->get_MacroName());

        ASSERT_EQ(u" MACROBUTTON  ViewZoom200 Run ViewZoom200", field->GetFieldCode());

        // Save the document as a macro-enabled document type
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

        // Add some keywords, also referred to as "tags" in File Explorer
        doc->get_BuiltInDocumentProperties()->set_Keywords(u"Keyword1, Keyword2");

        // Add a KEYWORDS field which will display our keywords
        auto field = System::DynamicCast<FieldKeywords>(builder->InsertField(FieldType::FieldKeyword, true));
        field->Update();

        ASSERT_EQ(u" KEYWORDS ", field->GetFieldCode());
        ASSERT_EQ(u"Keyword1, Keyword2", field->get_Result());

        // We can set the Text property of our field to display a different value to the one within the document's properties
        field->set_Text(u"OverridingKeyword");
        field->Update();

        ASSERT_EQ(u" KEYWORDS  OverridingKeyword", field->GetFieldCode());
        ASSERT_EQ(u"OverridingKeyword", field->get_Result());

        // Setting a KEYWORDS field's Text property also updates the document's keywords to our new value
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
        // Open a document to which we want to add character/word/page counts
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Move the document builder to the footer, where we will store our fields
        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);

        // Insert character and word counts
        auto fieldNumChars = System::DynamicCast<FieldNumChars>(builder->InsertField(FieldType::FieldNumChars, true));
        builder->Writeln(u" characters");
        auto fieldNumWords = System::DynamicCast<FieldNumWords>(builder->InsertField(FieldType::FieldNumWords, true));
        builder->Writeln(u" words");

        // Insert a "Page x of y" page count
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);
        builder->Write(u"Page ");
        auto fieldPage = System::DynamicCast<FieldPage>(builder->InsertField(FieldType::FieldPage, true));
        builder->Write(u" of ");
        auto fieldNumPages = System::DynamicCast<FieldNumPages>(builder->InsertField(FieldType::FieldNumPages, true));

        ASSERT_EQ(u" NUMCHARS ", fieldNumChars->GetFieldCode());
        ASSERT_EQ(u" NUMWORDS ", fieldNumWords->GetFieldCode());
        ASSERT_EQ(u" NUMPAGES ", fieldNumPages->GetFieldCode());
        ASSERT_EQ(u" PAGE ", fieldPage->GetFieldCode());

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

        // The PRINT field can send instructions to the printer that we use to print our document
        auto field = System::DynamicCast<FieldPrint>(builder->InsertField(FieldType::FieldPrint, true));

        // Set the area for the printer to perform instructions over
        // In this case, it will be the paragraph that contains our PRINT field
        field->set_PostScriptGroup(u"para");

        // When our document is printed using a printer that supports PostScript,
        // this command will turn the entire area that we specified in field.PostScriptGroup white
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

        // A PRINTDATE field will display "0/0/0000" by default
        // When a document is printed by a printer or printed as a PDF (but not exported as PDF),
        // these fields will display the date/time of that print operation
        auto field = System::DynamicCast<FieldPrintDate>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(u"3/25/2020 12:00:00 AM", field->get_Result());
        ASSERT_EQ(u" PRINTDATE ", field->GetFieldCode());

        // These fields can also display the date using other various international calendars
        field = System::DynamicCast<FieldPrintDate>(doc->get_Range()->get_Fields()->idx_get(1));

        ASSERT_TRUE(field->get_UseLunarCalendar());
        ASSERT_EQ(u"8/1/1441 12:00:00 AM", field->get_Result());
        ASSERT_EQ(u" PRINTDATE  \\h", field->GetFieldCode());

        field = System::DynamicCast<FieldPrintDate>(doc->get_Range()->get_Fields()->idx_get(2));

        ASSERT_TRUE(field->get_UseUmAlQuraCalendar());
        ASSERT_EQ(u"8/1/1441 12:00:00 AM", field->get_Result());
        ASSERT_EQ(u" PRINTDATE  \\u", field->GetFieldCode());

        field = System::DynamicCast<FieldPrintDate>(doc->get_Range()->get_Fields()->idx_get(3));

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

        // Insert a QUOTE field, which will display content from the Text attribute
        auto field = System::DynamicCast<FieldQuote>(builder->InsertField(FieldType::FieldQuote, true));
        field->set_Text(u"\"Quoted text\"");

        ASSERT_EQ(u" QUOTE  \"\\\"Quoted text\\\"\"", field->GetFieldCode());

        // Insert a QUOTE field with a nested DATE field
        // DATE fields normally update their value to the current date every time the document is opened
        // Nesting the DATE field inside the QUOTE field like this will freeze its value to the date when we created the document
        builder->Write(u"\nDocument creation date: ");
        field = System::DynamicCast<FieldQuote>(builder->InsertField(FieldType::FieldQuote, true));
        builder->MoveTo(field->get_Separator());
        builder->InsertField(FieldType::FieldDate, true);

        ASSERT_EQ(String(u" QUOTE \u0013 DATE \u0014") + System::DateTime::get_Now().get_Date().ToShortDateString() + u"\u0015", field->GetFieldCode());

        // Some field types don't display the correct result until they are manually updated
        ASSERT_EQ(String::Empty, doc->get_Range()->get_Fields()->idx_get(0)->get_Result());

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
    //ExSummary:Shows to insert NOTEREF fields and modify their appearance.
    void FieldNoteRef_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a bookmark with a footnote for the NOTEREF field to reference
        InsertBookmarkWithFootnote(builder, u"MyBookmark1", u"Contents of MyBookmark1", u"Footnote from MyBookmark1");

        // This NOTEREF field will display just the number of the footnote inside the referenced bookmark
        // Setting the InsertHyperlink attribute lets us jump to the bookmark by Ctrl + clicking the field
        ASSERT_EQ(u" NOTEREF  MyBookmark2 \\h",
                  InsertFieldNoteRef(builder, u"MyBookmark2", true, false, false, u"Hyperlink to Bookmark2, with footnote number ")->GetFieldCode());

        // When using the \p flag, after the footnote number the field also displays the position of the bookmark relative to the field
        // Bookmark1 is above this field and contains footnote number 1, so the result will be "1 above" on update
        ASSERT_EQ(u" NOTEREF  MyBookmark1 \\h \\p",
                  InsertFieldNoteRef(builder, u"MyBookmark1", true, true, false, u"Bookmark1, with footnote number ")->GetFieldCode());

        // Bookmark2 is below this field and contains footnote number 2, so the field will display "2 below"
        // The \f flag makes the number 2 appear in the same format as the footnote number label in the actual text
        ASSERT_EQ(u" NOTEREF  MyBookmark2 \\h \\p \\f",
                  InsertFieldNoteRef(builder, u"MyBookmark2", true, true, true, u"Bookmark2, with footnote number ")->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);
        InsertBookmarkWithFootnote(builder, u"MyBookmark2", u"Contents of MyBookmark2", u"Footnote from MyBookmark2");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.NOTEREF.docx");
        TestNoteRef(MakeObject<Document>(ArtifactsDir + u"Field.NOTEREF.docx"));
        //ExSkip
    }

protected:
    /// <summary>
    /// Uses a document builder to insert a NOTEREF field and sets its attributes.
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
    static void InsertBookmarkWithFootnote(SharedPtr<DocumentBuilder> builder, String bookmarkName, String bookmarkText,
                                           String footnoteText)
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

public:
    void FootnoteRef()
    {
        //ExStart
        //ExFor:FieldFootnoteRef
        //ExSummary:Shows how to cross-reference footnotes with the FOOTNOTEREF field
        // Create a blank document and a document builder for it
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert some text, and a footnote, all inside a bookmark named "CrossRefBookmark"
        builder->StartBookmark(u"CrossRefBookmark");
        builder->Write(u"Hello world!");
        builder->InsertFootnote(FootnoteType::Footnote, u"Cross referenced footnote.");
        builder->EndBookmark(u"CrossRefBookmark");

        builder->InsertParagraph();
        builder->Write(u"CrossReference: ");

        // Insert a FOOTNOTEREF field, which lets us reference a footnote more than once while re-using the same footnote marker
        auto field = System::DynamicCast<FieldFootnoteRef>(builder->InsertField(FieldType::FieldFootnoteRef, true));

        // Get this field to reference a bookmark
        // The bookmark that we chose contains a footnote marker belonging to the footnote we inserted, which will be displayed by the field, just by itself
        builder->MoveTo(field->get_Separator());
        builder->Write(u"CrossRefBookmark");

        ASSERT_EQ(u" FOOTNOTEREF CrossRefBookmark", field->GetFieldCode());

        doc->UpdateFields();

        // This field works only in older versions of Microsoft Word
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
    //ExSummary:Shows to insert PAGEREF fields and present them in different ways.
    void FieldPageRef_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        InsertAndNameBookmark(builder, u"MyBookmark1");

        // This field will display just the page number where the bookmark starts
        // Setting InsertHyperlink attribute makes the field function as a link to the bookmark
        ASSERT_EQ(u" PAGEREF  MyBookmark3 \\h", InsertFieldPageRef(builder, u"MyBookmark3", true, false, u"Hyperlink to Bookmark3, on page: ")->GetFieldCode());

        // Setting the \p flag makes the field display the relative position of the bookmark to the field instead of a page number
        // Bookmark1 is on the same page and above this field, so the result will be "above" on update
        ASSERT_EQ(u" PAGEREF  MyBookmark1 \\h \\p", InsertFieldPageRef(builder, u"MyBookmark1", true, true, u"Bookmark1 is ")->GetFieldCode());

        // Bookmark2 will be on the same page and below this field, so the field will display "below"
        ASSERT_EQ(u" PAGEREF  MyBookmark2 \\h \\p", InsertFieldPageRef(builder, u"MyBookmark2", true, true, u"Bookmark2 is ")->GetFieldCode());

        // Bookmark3 will be on a different page, so the field will display "on page 2"
        ASSERT_EQ(u" PAGEREF  MyBookmark3 \\h \\p", InsertFieldPageRef(builder, u"MyBookmark3", true, true, u"Bookmark3 is ")->GetFieldCode());

        InsertAndNameBookmark(builder, u"MyBookmark2");
        builder->InsertBreak(BreakType::PageBreak);
        InsertAndNameBookmark(builder, u"MyBookmark3");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.PAGEREF.docx");
        TestPageRef(MakeObject<Document>(ArtifactsDir + u"Field.PAGEREF.docx"));
        //ExSkip
    }

protected:
    /// <summary>
    /// Uses a document builder to insert a PAGEREF field and sets its attributes.
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

public:
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
    //ExSummary:Shows how to insert REF fields to reference bookmarks and present them in various ways.
    void FieldRef_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert the bookmark that all our REF fields will reference and leave it at the end of the document
        builder->StartBookmark(u"MyBookmark");
        builder->InsertFootnote(FootnoteType::Footnote, u"MyBookmark footnote #1");
        builder->Write(u"Text that will appear in REF field");
        builder->InsertFootnote(FootnoteType::Footnote, u"MyBookmark footnote #2");
        builder->EndBookmark(u"MyBookmark");
        builder->MoveToDocumentStart();

        // We will apply a custom list format, where the amount of angle brackets indicates the list level we are currently at
        // Note that the angle brackets count as non-delimiter characters
        builder->get_ListFormat()->ApplyNumberDefault();
        builder->get_ListFormat()->get_ListLevel()->set_NumberFormat(u"> \x0000");

        // Insert a REF field that will contain the text within our bookmark, act as a hyperlink, and clone the bookmark's footnotes
        SharedPtr<FieldRef> field = InsertFieldRef(builder, u"MyBookmark", u"", u"\n");
        field->set_IncludeNoteOrComment(true);
        field->set_InsertHyperlink(true);

        ASSERT_EQ(u" REF  MyBookmark \\f \\h", field->GetFieldCode());

        // Insert a REF field and display whether the referenced bookmark is above or below it
        field = InsertFieldRef(builder, u"MyBookmark", u"The referenced paragraph is ", u" this field.\n");
        field->set_InsertRelativePosition(true);

        ASSERT_EQ(u" REF  MyBookmark \\p", field->GetFieldCode());

        // Display the list number of the bookmark, as it appears in the document
        field = InsertFieldRef(builder, u"MyBookmark", u"The bookmark's paragraph number is ", u"\n");
        field->set_InsertParagraphNumber(true);

        ASSERT_EQ(u" REF  MyBookmark \\n", field->GetFieldCode());

        // Display the list number of the bookmark, but with non-delimiter characters omitted
        // In this case they are the angle brackets
        field = InsertFieldRef(builder, u"MyBookmark", u"The bookmark's paragraph number, non-delimiters suppressed, is ", u"\n");
        field->set_InsertParagraphNumber(true);
        field->set_SuppressNonDelimiters(true);

        ASSERT_EQ(u" REF  MyBookmark \\n \\t", field->GetFieldCode());

        // Move down one list level
        builder->get_ListFormat()->set_ListLevelNumber(builder->get_ListFormat()->get_ListLevelNumber() + 1);
        builder->get_ListFormat()->get_ListLevel()->set_NumberFormat(u">> \x0001");

        // Display the list number of the bookmark as well as the numbers of all the list levels above it
        field = InsertFieldRef(builder, u"MyBookmark", u"The bookmark's full context paragraph number is ", u"\n");
        field->set_InsertParagraphNumberInFullContext(true);

        ASSERT_EQ(u" REF  MyBookmark \\w", field->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);

        // Display the list level numbers between this REF field and the bookmark that it is referencing
        field = InsertFieldRef(builder, u"MyBookmark", u"The bookmark's relative paragraph number is ", u"\n");
        field->set_InsertParagraphNumberInRelativeContext(true);

        ASSERT_EQ(u" REF  MyBookmark \\r", field->GetFieldCode());

        // The bookmark, which is at the end of the document, will show up as a list item here
        builder->Writeln(u"List level above bookmark");
        builder->get_ListFormat()->set_ListLevelNumber(builder->get_ListFormat()->get_ListLevelNumber() + 1);
        builder->get_ListFormat()->get_ListLevel()->set_NumberFormat(u">>> \x0002");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.REF.docx");
        TestFieldRef(MakeObject<Document>(ArtifactsDir + u"Field.REF.docx"));
        //ExSkip
    }

protected:
    /// <summary>
    /// Get the document builder to insert a REF field, reference a bookmark with it, and add text before and after.
    /// </summary>
    static SharedPtr<FieldRef> InsertFieldRef(SharedPtr<DocumentBuilder> builder, String bookmarkName, String textBefore,
                                                      String textAfter)
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

public:
    void FieldRD_()
    {
        //ExStart
        //ExFor:FieldRD
        //ExFor:FieldRD.FileName
        //ExFor:FieldRD.IsPathRelative
        //ExSummary:Shows to insert an RD field to source table of contents entries from an external document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to insert a table of contents and, on the following page, one entry
        builder->InsertField(FieldType::FieldTOC, true);
        builder->InsertBreak(BreakType::PageBreak);
        builder->get_CurrentParagraph()->get_ParagraphFormat()->set_StyleName(u"Heading 1");
        builder->Writeln(u"TOC entry from within this document");

        // Insert an RD field, designating an external document that our TOC field will look in for more entries
        auto field = System::DynamicCast<FieldRD>(builder->InsertField(FieldType::FieldRefDoc, true));
        field->set_FileName(u"ReferencedDocument.docx");
        field->set_IsPathRelative(true);
        field->Update();

        ASSERT_EQ(u" RD  ReferencedDocument.docx \\f", field->GetFieldCode());

        // Create the document and insert a TOC entry, which will end up in the TOC of our original document
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

        ASSERT_EQ(String(u"TOC entry from within this document\t\u0013 PAGEREF _Toc36149519 \\h \u00142\u0015\r") +
                      u"TOC entry from referenced document\t1\r",
                  fieldToc->get_Result());

        auto fieldPageRef = System::DynamicCast<FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(1));

        TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF _Toc36149519 \\h ", u"2", fieldPageRef);

        field = System::DynamicCast<FieldRD>(doc->get_Range()->get_Fields()->idx_get(2));

        TestUtil::VerifyField(FieldType::FieldRefDoc, u" RD  ReferencedDocument.docx \\f", String::Empty, field);
        ASSERT_EQ(u"ReferencedDocument.docx", field->get_FileName());
        ASSERT_TRUE(field->get_IsPathRelative());
    }

    void FieldSet_()
    {
        //ExStart
        //ExFor:FieldSet
        //ExFor:FieldSet.BookmarkName
        //ExFor:FieldSet.BookmarkText
        //ExSummary:Shows to alter a bookmark's text with a SET field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartBookmark(u"MyBookmark");
        builder->Writeln(u"Bookmark contents");
        builder->EndBookmark(u"MyBookmark");

        SharedPtr<Bookmark> bookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"MyBookmark");
        bookmark->set_Text(u"Old text");

        auto field = System::DynamicCast<FieldSet>(builder->InsertField(FieldType::FieldSet, false));
        field->set_BookmarkName(u"MyBookmark");
        field->set_BookmarkText(u"New text");

        ASSERT_EQ(u" SET  MyBookmark \"New text\"", field->GetFieldCode());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"Field.SET.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.SET.docx");

        ASSERT_EQ(u"New text", doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Text());

        field = System::DynamicCast<FieldSet>(doc->get_Range()->get_Fields()->idx_get(0));

        TestUtil::VerifyField(FieldType::FieldSet, u" SET  MyBookmark \"New text\"", u"New text", field);
        ASSERT_EQ(u"MyBookmark", field->get_BookmarkName());
        ASSERT_EQ(u"New text", field->get_BookmarkText());
    }

    void FieldTemplate_()
    {
        //ExStart
        //ExFor:FieldTemplate
        //ExFor:FieldTemplate.IncludeFullPath
        //ExSummary:Shows how to display the location of the document's template with a TEMPLATE field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

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
        ASSERT_TRUE(field->get_Result().EndsWith(u"\\Microsoft\\Templates\\Normal.dotm"));
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

        // Insert a SYMBOL field to display a symbol, designated by a character code
        auto field = System::DynamicCast<FieldSymbol>(builder->InsertField(FieldType::FieldSymbol, true));

        // The ANSI character code "U+00A9", or "169" in integer form, is reserved for the copyright symbol
        field->set_CharacterCode(System::Convert::ToString(0x00a9));
        field->set_IsAnsi(true);

        ASSERT_EQ(u" SYMBOL  169 \\a", field->GetFieldCode());

        builder->Writeln(u" Line 1");

        // In Unicode, the "221E" code is reserved for the infinity symbol
        field = System::DynamicCast<FieldSymbol>(builder->InsertField(FieldType::FieldSymbol, true));
        field->set_CharacterCode(System::Convert::ToString(0x221E));
        field->set_IsUnicode(true);

        // Change the appearance of our symbol
        // Note that some symbols can change from font to font
        // The full list of symbols and their fonts can be looked up in the Windows Character Map
        field->set_FontName(u"Calibri");
        field->set_FontSize(u"24");

        // A tall symbol like the one we placed can also be made to not push down the text on its line
        field->set_DontAffectsLineSpacing(true);

        ASSERT_EQ(u" SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", field->GetFieldCode());

        builder->Writeln(u"Line 2");

        // Display a symbol from the Shift-JIS, also known as the Windows-932 code page
        // With a font that supports Shift-JIS, this symbol will display "あ"
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

        // A TITLE field will display the value assigned to this variable
        doc->get_BuiltInDocumentProperties()->set_Title(u"My Title");

        // Insert a TITLE field using a document builder
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto field = System::DynamicCast<FieldTitle>(builder->InsertField(FieldType::FieldTitle, false));
        field->Update();

        ASSERT_EQ(u" TITLE ", field->GetFieldCode());
        ASSERT_EQ(u"My Title", field->get_Result());

        // Set the Text attribute to display a different value
        builder->Writeln();
        field = System::DynamicCast<FieldTitle>(builder->InsertField(FieldType::FieldTitle, false));
        field->set_Text(u"My New Title");
        field->Update();

        ASSERT_EQ(u" TITLE  \"My New Title\"", field->GetFieldCode());
        ASSERT_EQ(u"My New Title", field->get_Result());

        // In doing that we have also changed the title in the document properties
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

        // Insert a TOA field, which will list all the TA entries in the document,
        // displaying long citations and page numbers for each
        auto fieldToa = System::DynamicCast<FieldToa>(builder->InsertField(FieldType::FieldTOA, false));

        // Set the entry category for our table
        // For a TA field to be included in this table, it will have to have a matching entry category
        fieldToa->set_EntryCategory(u"1");

        // Moreover, the Table of Authorities category at index 1 is "Cases",
        // which will show up as the title of our table if we set this variable to true
        fieldToa->set_UseHeading(true);

        // We can further filter TA fields by designating a named bookmark that they have to be inside of
        fieldToa->set_BookmarkName(u"MyBookmark");

        // By default, a dotted line page-wide tab appears between the TA field's citation and its page number
        // We can replace it with any text we put in this attribute, even preserving the tab if we use tab character
        fieldToa->set_EntrySeparator(u" \t p.");

        // If we have multiple TA entries that share the same long citation,
        // all their respective page numbers will show up on one row,
        // and the page numbers separated by a string specified here
        fieldToa->set_PageNumberListSeparator(u" & p. ");

        // To reduce clutter, we can set this to true to get our table to display the word "passim"
        // if there are 5 or more page numbers in one row
        fieldToa->set_UsePassim(true);

        // One TA field can refer to a range of pages, and the sequence specified here will be between the start and end page numbers
        fieldToa->set_PageRangeSeparator(u" to ");

        // The format from the TA fields will carry over into our table, and we can stop it from doing so by setting this variable
        fieldToa->set_RemoveEntryFormatting(true);
        builder->get_Font()->set_Color(System::Drawing::Color::get_Green());
        builder->get_Font()->set_Name(u"Arial Black");

        ASSERT_EQ(u" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f", fieldToa->GetFieldCode());

        builder->InsertBreak(BreakType::PageBreak);

        // We will insert a TA entry using a document builder
        // This entry is outside the bookmark specified by our table, so it will not be displayed
        SharedPtr<FieldTA> fieldTA = InsertToaEntry(builder, u"1", u"Source 1");

        ASSERT_EQ(u" TA  \\c 1 \\l \"Source 1\"", fieldTA->GetFieldCode());

        // This entry is inside the bookmark,
        // but the entry category does not match that of the table, so it will also be omitted
        builder->StartBookmark(u"MyBookmark");
        fieldTA = InsertToaEntry(builder, u"2", u"Source 2");

        // This entry will appear in the table
        fieldTA = InsertToaEntry(builder, u"1", u"Source 3");

        // Short citations are not displayed by a TOA table,
        // but they can be used as a shorthand to refer to bulky source names that multiple TA fields reference
        fieldTA->set_ShortCitation(u"S.3");

        ASSERT_EQ(u" TA  \\c 1 \\l \"Source 3\" \\s S.3", fieldTA->GetFieldCode());

        // The page number can be made to appear bold and/or italic
        // This will still be displayed if our table is set to ignore formatting
        fieldTA = InsertToaEntry(builder, u"1", u"Source 2");
        fieldTA->set_IsBold(true);
        fieldTA->set_IsItalic(true);

        ASSERT_EQ(u" TA  \\c 1 \\l \"Source 2\" \\b \\i", fieldTA->GetFieldCode());

        // We can get TA fields to refer to a range of pages that a bookmark spans across instead of the page that they are on
        // Note that this entry refers to the same source as the one above, so they will share one row in our table,
        // displaying the page number of the entry above as well as the page range of this entry,
        // with the table's page list and page number range separators between page numbers
        fieldTA = InsertToaEntry(builder, u"1", u"Source 3");
        fieldTA->set_PageRangeBookmarkName(u"MyMultiPageBookmark");

        builder->StartBookmark(u"MyMultiPageBookmark");
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::PageBreak);
        builder->EndBookmark(u"MyMultiPageBookmark");

        ASSERT_EQ(u" TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", fieldTA->GetFieldCode());

        // Having 5 or more TA entries with the same source invokes the "passim" feature of our table, if we enabled it
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

protected:
    /// <summary>
    /// Get a builder to insert a TA field, specifying its long citation and category,
    /// then insert a page break and return the field we created.
    /// </summary>
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

public:
    void FieldAddIn_()
    {
        //ExStart
        //ExFor:FieldAddIn
        //ExSummary:Shows how to process an ADDIN field.
        // Open a document that contains an ADDIN field
        auto doc = MakeObject<Document>(MyDir + u"Field sample - ADDIN.docx");

        // Aspose.Words does not support inserting ADDIN fields, they can be read
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

        // Use a document builder to insert an EDITTIME field in the header
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Write(u"You've been editing this document for ");
        auto field = System::DynamicCast<FieldEditTime>(builder->InsertField(FieldType::FieldEditTime, true));
        builder->Writeln(u" minutes.");

        // The EDITTIME field will show, in minutes only,
        // the time spent with the document open in a Microsoft Word window
        // The minutes are tracked in a document property, which we can change like this
        doc->get_BuiltInDocumentProperties()->set_TotalEditingTime(10);
        field->Update();

        ASSERT_EQ(u" EDITTIME ", field->GetFieldCode());
        ASSERT_EQ(u"10", field->get_Result());

        // The field does not update in real time and will have to be manually updated in Microsoft Word also
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

        // An EQ field displays a mathematical equation consisting of one or many elements
        // Each element takes the following form: [switch][options][arguments]
        // One switch, several possible options, followed by a set of argument values inside round braces

        // Here we use a document builder to insert an EQ field, with an "\f" switch, which corresponds to "Fraction"
        // No options are invoked, and the values 1 and 4 are passed as arguments
        // This field will display a fraction with 1 as the numerator and 4 as the denominator
        SharedPtr<FieldEQ> field = InsertFieldEQ(builder, u"\\f(1,4)");

        ASSERT_EQ(u" EQ \\f(1,4)", field->GetFieldCode());

        // One EQ field may contain multiple elements placed sequentially,
        // and elements may also be nested by being placed inside the argument brackets of other elements
        // The full list of switches and their corresponding options can be found here:
        // https://blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/

        // Array switch "\a", aligned left, 2 columns, 3 points of horizontal and vertical spacing
        InsertFieldEQ(builder, u"\\a \\al \\co2 \\vs3 \\hs3(4x,- 4y,-4x,+ y)");

        // Bracket switch "\b", bracket character "[", to enclose the contents in a set of square braces
        // Note that we are nesting an array inside the brackets, which will altogether look like a matrix in the output
        InsertFieldEQ(builder, u"\\b \\bc\\[ (\\a \\al \\co3 \\vs3 \\hs3(1,0,0,0,1,0,0,0,1))");

        // Displacement switch "\d", displacing text "B" 30 spaces to the right of "A", displaying the gap as an underline
        InsertFieldEQ(builder, u"A \\d \\fo30 \\li() B");

        // Formula consisting of multiple fractions
        InsertFieldEQ(builder, u"\\f(d,dx)(u + v) = \\f(du,dx) + \\f(dv,dx)");

        // Integral switch "\i", with a summation symbol
        InsertFieldEQ(builder, u"\\i \\su(n=1,5,n)");

        // List switch "\l"
        InsertFieldEQ(builder, u"\\l(1,1,2,3,n,8,13)");

        // Radical switch "\r", displaying a cubed root of x
        InsertFieldEQ(builder, u"\\r (3,x)");

        // Subscript/superscript switch "/s", first as a superscript and then as a subscript
        InsertFieldEQ(builder, u"\\s \\up8(Superscript) Text \\s \\do8(Subscript)");

        // Box switch "\x", with lines at the top, bottom, left and right of the input
        InsertFieldEQ(builder, u"\\x \\to \\bo \\le \\ri(5)");

        // More complex combinations
        InsertFieldEQ(builder, u"\\a \\ac \\vs1 \\co1(lim,n→∞) \\b (\\f(n,n2 + 12) + \\f(n,n2 + 22) + ... + \\f(n,n2 + n2))");
        InsertFieldEQ(builder, u"\\i (,,  \\b(\\f(x,x2 + 3x + 2))) \\s \\up10(2)");
        InsertFieldEQ(builder, u"\\i \\in( tan x, \\s \\up2(sec x), \\b(\\r(3) )\\s \\up4(t) \\s \\up7(2)  dt)");

        doc->Save(ArtifactsDir + u"Field.EQ.docx");
        TestFieldEQ(MakeObject<Document>(ArtifactsDir + u"Field.EQ.docx"));
        //ExSkip
    }

protected:
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
        TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\i \\in( tan x, \\s \\up2(sec x), \\b(\\r(3) )\\s \\up4(t) \\s \\up7(2)  dt)",
                              String::Empty, doc->get_Range()->get_Fields()->idx_get(12));
    }

public:
    void FieldForms()
    {
        //ExStart
        //ExFor:FieldFormCheckBox
        //ExFor:FieldFormDropDown
        //ExFor:FieldFormText
        //ExSummary:Shows how to process FORMCHECKBOX, FORMDROPDOWN and FORMTEXT fields.
        // These fields are legacy equivalents of the FormField, and they can be read but not inserted by Aspose.Words,
        // and can be inserted in Microsoft Word 2019 via the Legacy Tools menu in the Developer tab
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
        //ExSummary:Shows how to use the "=" field.
        auto doc = MakeObject<Document>();

        // Create a formula field using a field builder
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

        // If we create a document in Microsoft Word, it will have the user's name in the "Last saved by" property
        // This is the property that a LASTSAVEDBY field looks up and displays
        // If we make a document programmatically, this property is null and needs to have a value assigned to it first
        doc->get_BuiltInDocumentProperties()->set_LastSavedBy(u"John Doe");

        // Insert a LASTSAVEDBY field using a document builder
        auto field = System::DynamicCast<FieldLastSavedBy>(builder->InsertField(FieldType::FieldLastSavedBy, true));

        // The value from our document property appears here
        ASSERT_EQ(u" LASTSAVEDBY ", field->GetFieldCode());
        ASSERT_EQ(u"John Doe", field->get_Result());

        doc->UpdateFields();
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

        // Use a document builder to insert an OCX field
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
        // Open a Corel WordPerfect document that was converted to .docx format
        auto doc = MakeObject<Document>(MyDir + u"Field sample - PRIVATE.docx");

        // WordPerfect 5.x/6.x documents like the one we opened may contain PRIVATE fields
        // The PRIVATE field is a WordPerfect artifact that is preserved when a file is opened and saved in Microsoft Word
        // However, they have no functionality in Microsoft Word
        auto field = System::DynamicCast<FieldPrivate>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_EQ(u" PRIVATE \"My value\" ", field->GetFieldCode());
        ASSERT_EQ(FieldType::FieldPrivate, field->get_Type());

        // PRIVATE fields can also be inserted by a document builder
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertField(FieldType::FieldPrivate, true);

        // It is strongly advised against using them to attempt to hide or store private information
        // Unless backward compatibility with older versions of WordPerfect is necessary, these fields can safely be removed
        // This can be done using a document visitor implementation
        ASSERT_EQ(2, doc->get_Range()->get_Fields()->get_Count());

        auto remover = MakeObject<ExField::FieldPrivateRemover>();
        doc->Accept(remover);

        ASSERT_EQ(2, remover->GetFieldsRemovedCount());
        ASSERT_EQ(0, doc->get_Range()->get_Fields()->get_Count());
    }

    /// <summary>
    /// Visitor implementation that removes all PRIVATE fields that it encounters.
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
        //ExSummary:Shows how to use SECTION and SECTIONPAGES fields to facilitate page numbering separated by sections.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Move the document builder to the header that appears across all pages and align to the top right
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);

        // A SECTION field displays the number of the section it is placed in
        builder->Write(u"Section ");
        auto fieldSection = System::DynamicCast<FieldSection>(builder->InsertField(FieldType::FieldSection, true));

        ASSERT_EQ(u" SECTION ", fieldSection->GetFieldCode());

        // A PAGE field displays the number of the page it is placed in
        builder->Write(u"\nPage ");
        auto fieldPage = System::DynamicCast<FieldPage>(builder->InsertField(FieldType::FieldPage, true));

        ASSERT_EQ(u" PAGE ", fieldPage->GetFieldCode());

        // A SECTIONPAGES field displays the number of pages that the section it is in spans across
        builder->Write(u" of ");
        auto fieldSectionPages = System::DynamicCast<FieldSectionPages>(builder->InsertField(FieldType::FieldSectionPages, true));

        ASSERT_EQ(u" SECTIONPAGES ", fieldSectionPages->GetFieldCode());

        // Move out of the header back into the main document and insert two pages
        // Both these pages will be in the first section and our three fields will keep track of the numbers in each header
        builder->MoveToDocumentEnd();
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::PageBreak);

        // We can insert a new section with the document builder like this
        // This will change the values displayed in the SECTION and SECTIONPAGES fields in all upcoming headers
        builder->InsertBreak(BreakType::SectionBreakNewPage);

        // The PAGE field will keep counting pages across the whole document
        // We can manually reset its count after a new section is added to keep track of pages section-by-section
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

        // By default, time is displayed in the "h:mm am/pm" format
        SharedPtr<FieldTime> field = InsertFieldTime(builder, u"");

        ASSERT_EQ(u" TIME ", field->GetFieldCode());

        // By using the \@ flag, we can change the appearance of our time
        field = InsertFieldTime(builder, u"\\@ HHmm");

        ASSERT_EQ(u" TIME \\@ HHmm", field->GetFieldCode());

        // We can even display the date, according to the Gregorian calendar
        field = InsertFieldTime(builder, u"\\@ \"M/d/yyyy h mm:ss am/pm\"");

        ASSERT_EQ(u" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"", field->GetFieldCode());

        doc->Save(ArtifactsDir + u"Field.TIME.docx");
        TestFieldTime(MakeObject<Document>(ArtifactsDir + u"Field.TIME.docx"));
        //ExSkip
    }

protected:
    /// <summary>
    /// Use a document builder to insert a TIME field, insert a new paragraph and return the field
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

public:
    void BidiOutline()
    {
        //ExStart
        //ExFor:FieldBidiOutline
        //ExFor:FieldShape
        //ExFor:FieldShape.Text
        //ExFor:ParagraphFormat.Bidi
        //ExSummary:Shows how to create RTL lists with BIDIOUTLINE fields.
        // Create a blank document and a document builder
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use our builder to insert a BIDIOUTLINE field
        // This field numbers paragraphs like the AUTONUM/LISTNUM fields,
        // but is only visible when a RTL editing language is enabled, such as Hebrew or Arabic
        // The following field will display ".1", the RTL equivalent of list number "1."
        auto field = System::DynamicCast<FieldBidiOutline>(builder->InsertField(FieldType::FieldBidiOutline, true));
        builder->Writeln(u"שלום");

        ASSERT_EQ(u" BIDIOUTLINE ", field->GetFieldCode());

        // Add two more BIDIOUTLINE fields, which will be automatically numbered ".2" and ".3"
        builder->InsertField(FieldType::FieldBidiOutline, true);
        builder->Writeln(u"שלום");
        builder->InsertField(FieldType::FieldBidiOutline, true);
        builder->Writeln(u"שלום");

        // Set the horizontal text alignment for every paragraph in the document to RTL
        for (auto para : System::IterateOver<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)))
        {
            para->get_ParagraphFormat()->set_Bidi(true);
        }

        // If a RTL editing language is enabled in Microsoft Word, our fields will display numbers
        // Otherwise, they will appear as "###"
        doc->Save(ArtifactsDir + u"Field.BIDIOUTLINE.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Field.BIDIOUTLINE.docx");

        for (auto fieldBidiOutline : System::IterateOver(doc->get_Range()->get_Fields()))
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
        //ExSummary:Shows how some older Microsoft Word fields such as SHAPE and EMBED are handled.
        // Open a document that was created in Microsoft Word 2003
        auto doc = MakeObject<Document>(MyDir + u"Legacy fields.doc");

        // If we open the document in Word and press Alt+F9, we will see a SHAPE and an EMBED field
        // A SHAPE field is the anchor/canvas for an autoshape object with the "In line with text" wrapping style enabled
        // An EMBED field has the same function, but for an embedded object, such as a spreadsheet from an external Excel document
        // However, these fields will not appear in the document's Fields collection
        ASSERT_EQ(0, doc->get_Range()->get_Fields()->get_Count());

        // These fields are supported only by old versions of Microsoft Word
        // As such, they are converted into shapes during the document importation process and can instead be found in the collection of Shape nodes
        SharedPtr<NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);
        ASSERT_EQ(3, shapes->get_Count());

        // The first Shape node corresponds to what was the SHAPE field in the input document: the inline canvas for an autoshape
        auto shape = System::DynamicCast<Shape>(shapes->idx_get(0));
        ASSERT_EQ(ShapeType::Image, shape->get_ShapeType());

        // The next Shape node is the autoshape that is within the canvas
        shape = System::DynamicCast<Shape>(shapes->idx_get(1));
        ASSERT_EQ(ShapeType::Can, shape->get_ShapeType());

        // The third Shape is what was the EMBED field that contained the external spreadsheet
        shape = System::DynamicCast<Shape>(shapes->idx_get(2));
        ASSERT_EQ(ShapeType::OleObject, shape->get_ShapeType());
        //ExEnd
    }
};

} // namespace ApiExamples
