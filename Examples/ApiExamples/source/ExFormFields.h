#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Bookmark.h>
#include <Aspose.Words.Cpp/BookmarkCollection.h>
#include <Aspose.Words.Cpp/BookmarkEnd.h>
#include <Aspose.Words.Cpp/BookmarkStart.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Fields/DropDownItemCollection.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Fields/FormField.h>
#include <Aspose.Words.Cpp/Fields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Fields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/VisitorAction.h>
#include <drawing/color.h>
#include <system/array.h>
#include <system/collections/ienumerator.h>
#include <system/details/dispose_guard.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/string_builder.h>
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
using namespace Aspose::Words::Fields;

namespace ApiExamples {

class ExFormFields : public ApiExampleBase
{
public:
    void Create()
    {
        //ExStart
        //ExFor:FormField
        //ExFor:FormField.Result
        //ExFor:FormField.Type
        //ExFor:FormField.Name
        //ExSummary:Shows how to insert a combo box.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Please select a fruit: ");

        // Insert a combo box which will allow a user to choose an option from a collection of strings.
        SharedPtr<FormField> comboBox = builder->InsertComboBox(u"MyComboBox", MakeArray<String>({u"Apple", u"Banana", u"Cherry"}), 0);

        ASSERT_EQ(u"MyComboBox", comboBox->get_Name());
        ASSERT_EQ(FieldType::FieldFormDropDown, comboBox->get_Type());
        ASSERT_EQ(u"Apple", comboBox->get_Result());

        // The form field will appear in the form of a "select" html tag.
        doc->Save(ArtifactsDir + u"FormFields.Create.html");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"FormFields.Create.html");
        comboBox = doc->get_Range()->get_FormFields()->idx_get(0);

        ASSERT_EQ(u"MyComboBox", comboBox->get_Name());
        ASSERT_EQ(FieldType::FieldFormDropDown, comboBox->get_Type());
        ASSERT_EQ(u"Apple", comboBox->get_Result());
    }

    void TextInput()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertTextInput
        //ExSummary:Shows how to insert a text input form field.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Please enter text here: ");

        // Insert a text input field, which will allow the user to click it and enter text.
        // Assign some placeholder text that the user may overwrite and pass
        // a maximum text length of 0 to apply no limit on the form field's contents.
        builder->InsertTextInput(u"TextInput1", TextFormFieldType::Regular, u"", u"Placeholder text", 0);

        // The form field will appear in the form of an "input" html tag, with a type of "text".
        doc->Save(ArtifactsDir + u"FormFields.TextInput.html");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"FormFields.TextInput.html");

        SharedPtr<FormField> textInput = doc->get_Range()->get_FormFields()->idx_get(0);

        ASSERT_EQ(u"TextInput1", textInput->get_Name());
        ASSERT_EQ(TextFormFieldType::Regular, textInput->get_TextInputType());
        ASSERT_EQ(String::Empty, textInput->get_TextInputFormat());
        ASSERT_EQ(u"Placeholder text", textInput->get_Result());
        ASSERT_EQ(0, textInput->get_MaxLength());
    }

    void DeleteFormField()
    {
        //ExStart
        //ExFor:FormField.RemoveField
        //ExSummary:Shows how to delete a form field.
        auto doc = MakeObject<Document>(MyDir + u"Form fields.docx");

        SharedPtr<FormField> formField = doc->get_Range()->get_FormFields()->idx_get(3);
        formField->RemoveField();
        //ExEnd

        SharedPtr<FormField> formFieldAfter = doc->get_Range()->get_FormFields()->idx_get(3);

        ASSERT_TRUE(formFieldAfter == nullptr);
    }

    void DeleteFormFieldAssociatedWithBookmark()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartBookmark(u"MyBookmark");
        builder->InsertTextInput(u"TextInput1", TextFormFieldType::Regular, u"TestFormField", u"SomeText", 0);
        builder->EndBookmark(u"MyBookmark");

        doc = DocumentHelper::SaveOpen(doc);

        SharedPtr<BookmarkCollection> bookmarkBeforeDeleteFormField = doc->get_Range()->get_Bookmarks();
        ASSERT_EQ(u"MyBookmark", bookmarkBeforeDeleteFormField->idx_get(0)->get_Name());

        SharedPtr<FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);
        formField->RemoveField();

        SharedPtr<BookmarkCollection> bookmarkAfterDeleteFormField = doc->get_Range()->get_Bookmarks();
        ASSERT_EQ(u"MyBookmark", bookmarkAfterDeleteFormField->idx_get(0)->get_Name());
    }

    void FormFieldFontFormatting()
    {
        //ExStart
        //ExFor:FormField
        //ExSummary:Shows how to formatting the entire FormField, including the field value.
        auto doc = MakeObject<Document>(MyDir + u"Form fields.docx");

        SharedPtr<FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);
        formField->get_Font()->set_Bold(true);
        formField->get_Font()->set_Size(24);
        formField->get_Font()->set_Color(System::Drawing::Color::get_Red());

        formField->set_Result(u"Aspose.FormField");

        doc = DocumentHelper::SaveOpen(doc);

        SharedPtr<Run> formFieldRun = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(1);

        ASSERT_EQ(u"Aspose.FormField", formFieldRun->get_Text());
        ASPOSE_ASSERT_EQ(true, formFieldRun->get_Font()->get_Bold());
        ASPOSE_ASSERT_EQ(24, formFieldRun->get_Font()->get_Size());
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), formFieldRun->get_Font()->get_Color().ToArgb());
        //ExEnd
    }

    //ExStart
    //ExFor:FormField.Accept(DocumentVisitor)
    //ExFor:FormField.CalculateOnExit
    //ExFor:FormField.CheckBoxSize
    //ExFor:FormField.Checked
    //ExFor:FormField.Default
    //ExFor:FormField.DropDownItems
    //ExFor:FormField.DropDownSelectedIndex
    //ExFor:FormField.Enabled
    //ExFor:FormField.EntryMacro
    //ExFor:FormField.ExitMacro
    //ExFor:FormField.HelpText
    //ExFor:FormField.IsCheckBoxExactSize
    //ExFor:FormField.MaxLength
    //ExFor:FormField.OwnHelp
    //ExFor:FormField.OwnStatus
    //ExFor:FormField.SetTextInputValue(Object)
    //ExFor:FormField.StatusText
    //ExFor:FormField.TextInputDefault
    //ExFor:FormField.TextInputFormat
    //ExFor:FormField.TextInputType
    //ExFor:FormFieldCollection
    //ExFor:FormFieldCollection.Clear
    //ExFor:FormFieldCollection.Count
    //ExFor:FormFieldCollection.GetEnumerator
    //ExFor:FormFieldCollection.Item(Int32)
    //ExFor:FormFieldCollection.Item(String)
    //ExFor:FormFieldCollection.Remove(String)
    //ExFor:FormFieldCollection.RemoveAt(Int32)
    //ExFor:Range.FormFields
    //ExSummary:Shows how insert different kinds of form fields into a document, and process them with using a document visitor implementation.
    void Visitor()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to insert a combo box.
        builder->Write(u"Choose a value from this combo box: ");
        SharedPtr<FormField> comboBox = builder->InsertComboBox(u"MyComboBox", MakeArray<String>({u"One", u"Two", u"Three"}), 0);
        comboBox->set_CalculateOnExit(true);
        ASSERT_EQ(3, comboBox->get_DropDownItems()->get_Count());
        ASSERT_EQ(0, comboBox->get_DropDownSelectedIndex());
        ASSERT_TRUE(comboBox->get_Enabled());

        builder->InsertBreak(BreakType::ParagraphBreak);

        // Use a document builder to insert a check box.
        builder->Write(u"Click this check box to tick/untick it: ");
        SharedPtr<FormField> checkBox = builder->InsertCheckBox(u"MyCheckBox", false, 50);
        checkBox->set_IsCheckBoxExactSize(true);
        checkBox->set_HelpText(u"Right click to check this box");
        checkBox->set_OwnHelp(true);
        checkBox->set_StatusText(u"Checkbox status text");
        checkBox->set_OwnStatus(true);
        ASPOSE_ASSERT_EQ(50.0, checkBox->get_CheckBoxSize());
        ASSERT_FALSE(checkBox->get_Checked());
        ASSERT_FALSE(checkBox->get_Default());

        builder->InsertBreak(BreakType::ParagraphBreak);

        // Use a document builder to insert text input form field.
        builder->Write(u"Enter text here: ");
        SharedPtr<FormField> textInput = builder->InsertTextInput(u"MyTextInput", TextFormFieldType::Regular, u"", u"Placeholder text", 50);
        textInput->set_EntryMacro(u"EntryMacro");
        textInput->set_ExitMacro(u"ExitMacro");
        textInput->set_TextInputDefault(u"Regular");
        textInput->set_TextInputFormat(u"FIRST CAPITAL");
        textInput->SetTextInputValue(System::ObjectExt::Box<String>(u"New placeholder text"));
        ASSERT_EQ(TextFormFieldType::Regular, textInput->get_TextInputType());
        ASSERT_EQ(50, textInput->get_MaxLength());

        // This collection contains all our form fields.
        SharedPtr<FormFieldCollection> formFields = doc->get_Range()->get_FormFields();
        ASSERT_EQ(3, formFields->get_Count());

        // Fields display our form fields. We can see their field codes by opening this document
        // in Microsoft and pressing Alt + F9. These fields have no switches,
        // and members of the FormField object fully govern their form fields' content.
        ASSERT_EQ(3, doc->get_Range()->get_Fields()->get_Count());
        ASSERT_EQ(u" FORMDROPDOWN \u0001", doc->get_Range()->get_Fields()->idx_get(0)->GetFieldCode());
        ASSERT_EQ(u" FORMCHECKBOX \u0001", doc->get_Range()->get_Fields()->idx_get(1)->GetFieldCode());
        ASSERT_EQ(u" FORMTEXT \u0001", doc->get_Range()->get_Fields()->idx_get(2)->GetFieldCode());

        // Allow each form field to accept a document visitor.
        auto formFieldVisitor = MakeObject<ExFormFields::FormFieldVisitor>();

        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<FormField>>> fieldEnumerator = formFields->GetEnumerator();
            while (fieldEnumerator->MoveNext())
            {
                fieldEnumerator->get_Current()->Accept(formFieldVisitor);
            }
        }

        std::cout << formFieldVisitor->GetText() << std::endl;

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"FormFields.Visitor.html");
        TestFormField(doc);
        //ExSkip
    }

    /// <summary>
    /// Visitor implementation that prints details of form fields that it visits.
    /// </summary>
    class FormFieldVisitor : public DocumentVisitor
    {
    public:
        FormFieldVisitor()
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
        }

        /// <summary>
        /// Called when a FormField node is encountered in the document.
        /// </summary>
        VisitorAction VisitFormField(SharedPtr<FormField> formField) override
        {
            AppendLine(System::ObjectExt::ToString(formField->get_Type()) + u": \"" + formField->get_Name() + u"\"");
            AppendLine(String(u"\tStatus: ") + (formField->get_Enabled() ? String(u"Enabled") : String(u"Disabled")));
            AppendLine(String(u"\tHelp Text:  ") + formField->get_HelpText());
            AppendLine(String(u"\tEntry macro name: ") + formField->get_EntryMacro());
            AppendLine(String(u"\tExit macro name: ") + formField->get_ExitMacro());

            switch (formField->get_Type())
            {
            case FieldType::FieldFormDropDown:
                AppendLine(String(u"\tDrop-down items count: ") + formField->get_DropDownItems()->get_Count() + u", default selected item index: " +
                           formField->get_DropDownSelectedIndex());
                AppendLine(String(u"\tDrop-down items: ") + String::Join(u", ", formField->get_DropDownItems()->LINQ_ToArray()));
                break;

            case FieldType::FieldFormCheckBox:
                AppendLine(String(u"\tCheckbox size: ") + formField->get_CheckBoxSize());
                AppendLine(String(u"\t") + u"Checkbox is currently: " + (formField->get_Checked() ? String(u"checked, ") : String(u"unchecked, ")) +
                           u"by default: " + (formField->get_Default() ? String(u"checked") : String(u"unchecked")));
                break;

            case FieldType::FieldFormTextInput:
                AppendLine(String(u"\tInput format: ") + formField->get_TextInputFormat());
                AppendLine(String(u"\tCurrent contents: ") + formField->get_Result());
                break;

            default:
                break;
            }

            // Let the visitor continue visiting other nodes.
            return VisitorAction::Continue;
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        String GetText()
        {
            return mBuilder->ToString();
        }

    private:
        SharedPtr<System::Text::StringBuilder> mBuilder;

        /// <summary>
        /// Adds newline char-terminated text to the current output.
        /// </summary>
        void AppendLine(String text)
        {
            mBuilder->Append(text + u'\n');
        }
    };
    //ExEnd

    void TestFormField(SharedPtr<Document> doc)
    {
        doc = DocumentHelper::SaveOpen(doc);
        SharedPtr<FieldCollection> fields = doc->get_Range()->get_Fields();
        ASSERT_EQ(3, fields->get_Count());

        TestUtil::VerifyField(FieldType::FieldFormDropDown, u" FORMDROPDOWN \u0001", String::Empty, doc->get_Range()->get_Fields()->idx_get(0));
        TestUtil::VerifyField(FieldType::FieldFormCheckBox, u" FORMCHECKBOX \u0001", String::Empty, doc->get_Range()->get_Fields()->idx_get(1));
        TestUtil::VerifyField(FieldType::FieldFormTextInput, u" FORMTEXT \u0001", u"Regular", doc->get_Range()->get_Fields()->idx_get(2));

        SharedPtr<FormFieldCollection> formFields = doc->get_Range()->get_FormFields();
        ASSERT_EQ(3, formFields->get_Count());

        ASSERT_EQ(FieldType::FieldFormDropDown, formFields->idx_get(0)->get_Type());
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"One", u"Two", u"Three"}), formFields->idx_get(0)->get_DropDownItems());
        ASSERT_TRUE(formFields->idx_get(0)->get_CalculateOnExit());
        ASSERT_EQ(0, formFields->idx_get(0)->get_DropDownSelectedIndex());
        ASSERT_TRUE(formFields->idx_get(0)->get_Enabled());
        ASSERT_EQ(u"One", formFields->idx_get(0)->get_Result());

        ASSERT_EQ(FieldType::FieldFormCheckBox, formFields->idx_get(1)->get_Type());
        ASSERT_TRUE(formFields->idx_get(1)->get_IsCheckBoxExactSize());
        ASSERT_EQ(u"Right click to check this box", formFields->idx_get(1)->get_HelpText());
        ASSERT_TRUE(formFields->idx_get(1)->get_OwnHelp());
        ASSERT_EQ(u"Checkbox status text", formFields->idx_get(1)->get_StatusText());
        ASSERT_TRUE(formFields->idx_get(1)->get_OwnStatus());
        ASPOSE_ASSERT_EQ(50.0, formFields->idx_get(1)->get_CheckBoxSize());
        ASSERT_FALSE(formFields->idx_get(1)->get_Checked());
        ASSERT_FALSE(formFields->idx_get(1)->get_Default());
        ASSERT_EQ(u"0", formFields->idx_get(1)->get_Result());

        ASSERT_EQ(FieldType::FieldFormTextInput, formFields->idx_get(2)->get_Type());
        ASSERT_EQ(u"EntryMacro", formFields->idx_get(2)->get_EntryMacro());
        ASSERT_EQ(u"ExitMacro", formFields->idx_get(2)->get_ExitMacro());
        ASSERT_EQ(u"Regular", formFields->idx_get(2)->get_TextInputDefault());
        ASSERT_EQ(u"FIRST CAPITAL", formFields->idx_get(2)->get_TextInputFormat());
        ASSERT_EQ(TextFormFieldType::Regular, formFields->idx_get(2)->get_TextInputType());
        ASSERT_EQ(50, formFields->idx_get(2)->get_MaxLength());
        ASSERT_EQ(u"Regular", formFields->idx_get(2)->get_Result());
    }

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
        //ExSummary:Shows how to insert a combo box field, and edit the elements in its item collection.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a combo box, and then verify its collection of drop-down items.
        // In Microsoft Word, the user will click the combo box,
        // and then choose one of the items of text in the collection to display.
        ArrayPtr<String> items = MakeArray<String>({u"One", u"Two", u"Three"});
        SharedPtr<FormField> comboBoxField = builder->InsertComboBox(u"DropDown", items, 0);
        SharedPtr<DropDownItemCollection> dropDownItems = comboBoxField->get_DropDownItems();

        ASSERT_EQ(3, dropDownItems->get_Count());
        ASSERT_EQ(u"One", dropDownItems->idx_get(0));
        ASSERT_EQ(1, dropDownItems->IndexOf(u"Two"));
        ASSERT_TRUE(dropDownItems->Contains(u"Three"));

        // There are two ways of adding a new item to an existing collection of drop-down box items.
        // 1 -  Append an item to the end of the collection:
        dropDownItems->Add(u"Four");

        // 2 -  Insert an item before another item at a specified index:
        dropDownItems->Insert(3, u"Three and a half");

        ASSERT_EQ(5, dropDownItems->get_Count());

        // Iterate over the collection and print every element.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<String>> dropDownCollectionEnumerator = dropDownItems->GetEnumerator();
            while (dropDownCollectionEnumerator->MoveNext())
            {
                std::cout << dropDownCollectionEnumerator->get_Current() << std::endl;
            }
        }

        // There are two ways of removing elements from a collection of drop-down items.
        // 1 -  Remove an item with contents equal to the passed string:
        dropDownItems->Remove(u"Four");

        // 2 -  Remove an item at an index:
        dropDownItems->RemoveAt(3);

        ASSERT_EQ(3, dropDownItems->get_Count());
        ASSERT_FALSE(dropDownItems->Contains(u"Three and a half"));
        ASSERT_FALSE(dropDownItems->Contains(u"Four"));

        doc->Save(ArtifactsDir + u"FormFields.DropDownItemCollection.html");

        // Empty the whole collection of drop-down items.
        dropDownItems->Clear();
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        dropDownItems = doc->get_Range()->get_FormFields()->idx_get(0)->get_DropDownItems();

        ASSERT_EQ(0, dropDownItems->get_Count());

        doc = MakeObject<Document>(ArtifactsDir + u"FormFields.DropDownItemCollection.html");
        dropDownItems = doc->get_Range()->get_FormFields()->idx_get(0)->get_DropDownItems();

        ASSERT_EQ(3, dropDownItems->get_Count());
        ASSERT_EQ(u"One", dropDownItems->idx_get(0));
        ASSERT_EQ(u"Two", dropDownItems->idx_get(1));
        ASSERT_EQ(u"Three", dropDownItems->idx_get(2));
    }
};

} // namespace ApiExamples
