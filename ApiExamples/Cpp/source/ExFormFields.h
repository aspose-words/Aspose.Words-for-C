#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <iostream>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/DropDownItemCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
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

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace ApiExamples {

class ExFormFields : public ApiExampleBase
{
public:
    void FormFieldsWorkWithProperties()
    {
        //ExStart
        //ExFor:FormField
        //ExFor:FormField.Result
        //ExFor:FormField.Type
        //ExFor:FormField.Name
        //ExSummary:Shows how to work with form field name, type, and result.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a DocumentBuilder to insert a combo box form field
        SharedPtr<FormField> comboBox = builder->InsertComboBox(u"MyComboBox", MakeArray<String>({u"One", u"Two", u"Three"}), 0);

        // Verify some of our form field's attributes
        ASSERT_EQ(u"MyComboBox", comboBox->get_Name());
        ASSERT_EQ(FieldType::FieldFormDropDown, comboBox->get_Type());
        ASSERT_EQ(u"One", comboBox->get_Result());
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        comboBox = doc->get_Range()->get_FormFields()->idx_get(0);

        ASSERT_EQ(u"MyComboBox", comboBox->get_Name());
        ASSERT_EQ(FieldType::FieldFormDropDown, comboBox->get_Type());
        ASSERT_EQ(u"One", comboBox->get_Result());
    }

    void InsertAndRetrieveFormFields()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertTextInput
        //ExSummary:Shows how to insert form fields, set options and gather them back in for use.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a text input field. The unique name of this field is "TextInput1", the other parameters define
        // what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
        builder->InsertTextInput(u"TextInput1", TextFormFieldType::Regular, u"", u"", 0);
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        SharedPtr<FormField> textInput = doc->get_Range()->get_FormFields()->idx_get(0);

        ASSERT_EQ(u"TextInput1", textInput->get_Name());
        ASSERT_EQ(TextFormFieldType::Regular, textInput->get_TextInputType());
        ASSERT_EQ(String::Empty, textInput->get_TextInputFormat());
        ASSERT_EQ(String::Empty, textInput->get_Result());
        ASSERT_EQ(0, textInput->get_MaxLength());
    }

    void DeleteFormField()
    {
        //ExStart
        //ExFor:FormField.RemoveField
        //ExSummary:Shows how to delete complete form field.
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
    //ExSummary:Shows how insert different kinds of form fields into a document and process them with a visitor implementation.
    void FormField_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use a document builder to insert a combo box
        SharedPtr<FormField> comboBox = builder->InsertComboBox(u"MyComboBox", MakeArray<String>({u"One", u"Two", u"Three"}), 0);
        comboBox->set_CalculateOnExit(true);
        ASSERT_EQ(3, comboBox->get_DropDownItems()->get_Count());
        ASSERT_EQ(0, comboBox->get_DropDownSelectedIndex());
        ASSERT_TRUE(comboBox->get_Enabled());

        // Use a document builder to insert a check box
        SharedPtr<FormField> checkBox = builder->InsertCheckBox(u"MyCheckBox", false, 50);
        checkBox->set_IsCheckBoxExactSize(true);
        checkBox->set_HelpText(u"Right click to check this box");
        checkBox->set_OwnHelp(true);
        checkBox->set_StatusText(u"Checkbox status text");
        checkBox->set_OwnStatus(true);
        ASPOSE_ASSERT_EQ(50.0, checkBox->get_CheckBoxSize());
        ASSERT_FALSE(checkBox->get_Checked());
        ASSERT_FALSE(checkBox->get_Default());

        builder->Writeln();

        // Use a document builder to insert text input form field
        SharedPtr<FormField> textInput = builder->InsertTextInput(u"MyTextInput", TextFormFieldType::Regular, u"", u"Your text goes here", 50);
        textInput->set_EntryMacro(u"EntryMacro");
        textInput->set_ExitMacro(u"ExitMacro");
        textInput->set_TextInputDefault(u"Regular");
        textInput->set_TextInputFormat(u"FIRST CAPITAL");
        textInput->SetTextInputValue(System::ObjectExt::Box<String>(u"This value overrides the one we set during initialization"));
        ASSERT_EQ(TextFormFieldType::Regular, textInput->get_TextInputType());
        ASSERT_EQ(50, textInput->get_MaxLength());

        // Get the collection of form fields that has accumulated in our document
        SharedPtr<FormFieldCollection> formFields = doc->get_Range()->get_FormFields();
        ASSERT_EQ(3, formFields->get_Count());

        // Our form fields are represented as fields, with field codes FORMDROPDOWN, FORMCHECKBOX and FORMTEXT respectively,
        // made visible by pressing Alt + F9 in Microsoft Word
        // These fields have no switches and the content of their form fields is fully governed by members of the FormField object
        ASSERT_EQ(3, doc->get_Range()->get_Fields()->get_Count());

        // Iterate over the collection with an enumerator, accepting a visitor with each form field
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
        doc->Save(ArtifactsDir + u"Field.FormField.docx");
        TestFormField(doc);
        //ExSkip
    }

    /// <summary>
    /// Visitor implementation that prints information about visited form fields.
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
                AppendLine(String(u"\tDrop down items count: ") + formField->get_DropDownItems()->get_Count() + u", default selected item index: " +
                           formField->get_DropDownSelectedIndex());
                AppendLine(String(u"\tDrop down items: ") + String::Join(u", ", formField->get_DropDownItems()->LINQ_ToArray()));
                break;

            case FieldType::FieldFormCheckBox:
                AppendLine(String(u"\tCheckbox size: ") + formField->get_CheckBoxSize());
                AppendLine(String(u"\t") + u"Checkbox is currently: " +
                           (formField->get_Checked() ? String(u"checked, ") : String(u"unchecked, ")) + u"by default: " +
                           (formField->get_Default() ? String(u"checked") : String(u"unchecked")));
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

protected:
    void TestFormField(SharedPtr<Document> doc)
    {
        doc = DocumentHelper::SaveOpen(doc);
        SharedPtr<FieldCollection> fields = doc->get_Range()->get_Fields();
        ASSERT_EQ(3, fields->get_Count());

        TestUtil::VerifyField(FieldType::FieldFormDropDown, u" FORMDROPDOWN \u0001", String::Empty, doc->get_Range()->get_Fields()->idx_get(0));
        TestUtil::VerifyField(FieldType::FieldFormCheckBox, u" FORMCHECKBOX \u0001", String::Empty, doc->get_Range()->get_Fields()->idx_get(1));
        TestUtil::VerifyField(FieldType::FieldFormTextInput, u" FORMTEXT \u0001", u"This value overrides the one we set during initialization",
                              doc->get_Range()->get_Fields()->idx_get(2));

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
        ASSERT_EQ(u"This value overrides the one we set during initialization", formFields->idx_get(2)->get_Result());
    }
};

} // namespace ApiExamples
