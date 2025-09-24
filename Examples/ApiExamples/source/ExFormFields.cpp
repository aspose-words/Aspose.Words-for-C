// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExFormFields.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/details/dispose_guard.h>
#include <system/collections/ienumerator.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/DropDownItemCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>

#include "TestUtil.h"
#include "DocumentHelper.h"


using namespace Aspose::Words::Fields;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(417268581u, ::Aspose::Words::ApiExamples::ExFormFields::FormFieldVisitor, ThisTypeBaseTypesInfo);

ExFormFields::FormFieldVisitor::FormFieldVisitor()
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
}

Aspose::Words::VisitorAction ExFormFields::FormFieldVisitor::VisitFormField(System::SharedPtr<Aspose::Words::Fields::FormField> formField)
{
    AppendLine(System::ObjectExt::ToString(formField->get_Type()) + u": \"" + formField->get_Name() + u"\"");
    AppendLine(System::String(u"\tStatus: ") + (formField->get_Enabled() ? System::String(u"Enabled") : System::String(u"Disabled")));
    AppendLine(System::String(u"\tHelp Text:  ") + formField->get_HelpText());
    AppendLine(System::String(u"\tEntry macro name: ") + formField->get_EntryMacro());
    AppendLine(System::String(u"\tExit macro name: ") + formField->get_ExitMacro());
    
    switch (formField->get_Type())
    {
        case Aspose::Words::Fields::FieldType::FieldFormDropDown:
            AppendLine(System::String(u"\tDrop-down items count: ") + formField->get_DropDownItems()->get_Count() + u", default selected item index: " + formField->get_DropDownSelectedIndex());
            AppendLine(System::String(u"\tDrop-down items: ") + System::String::Join(u", ", formField->get_DropDownItems()->LINQ_ToArray()));
            break;
        
        case Aspose::Words::Fields::FieldType::FieldFormCheckBox:
            AppendLine(System::String(u"\tCheckbox size: ") + formField->get_CheckBoxSize());
            AppendLine(System::String(u"\t") + u"Checkbox is currently: " + (formField->get_Checked() ? System::String(u"checked, ") : System::String(u"unchecked, ")) + u"by default: " + (formField->get_Default() ? System::String(u"checked") : System::String(u"unchecked")));
            break;
        
        case Aspose::Words::Fields::FieldType::FieldFormTextInput:
            AppendLine(System::String(u"\tInput format: ") + formField->get_TextInputFormat());
            AppendLine(System::String(u"\tCurrent contents: ") + formField->get_Result());
            break;
        
        default:
            break;
    }
    
    // Let the visitor continue visiting other nodes.
    return Aspose::Words::VisitorAction::Continue;
}

void ExFormFields::FormFieldVisitor::AppendLine(System::String text)
{
    mBuilder->Append(text + u'\n');
}

System::String ExFormFields::FormFieldVisitor::GetText()
{
    return mBuilder->ToString();
}


RTTI_INFO_IMPL_HASH(1459817869u, ::Aspose::Words::ApiExamples::ExFormFields, ThisTypeBaseTypesInfo);

void ExFormFields::TestFormField(System::SharedPtr<Aspose::Words::Document> doc)
{
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    System::SharedPtr<Aspose::Words::Fields::FieldCollection> fields = doc->get_Range()->get_Fields();
    ASSERT_EQ(3, fields->get_Count());
    
    Aspose::Words::ApiExamples::TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldFormDropDown, u" FORMDROPDOWN \u0001", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldFormCheckBox, u" FORMCHECKBOX \u0001", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(1));
    Aspose::Words::ApiExamples::TestUtil::VerifyField(Aspose::Words::Fields::FieldType::FieldFormTextInput, u" FORMTEXT \u0001", u"Regular", doc->get_Range()->get_Fields()->idx_get(2));
    
    System::SharedPtr<Aspose::Words::Fields::FormFieldCollection> formFields = doc->get_Range()->get_FormFields();
    ASSERT_EQ(3, formFields->get_Count());
    
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormDropDown, formFields->idx_get(0)->get_Type());
    ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"One", u"Two", u"Three"}), formFields->idx_get(0)->get_DropDownItems());
    ASSERT_TRUE(formFields->idx_get(0)->get_CalculateOnExit());
    ASSERT_EQ(0, formFields->idx_get(0)->get_DropDownSelectedIndex());
    ASSERT_TRUE(formFields->idx_get(0)->get_Enabled());
    ASSERT_EQ(u"One", formFields->idx_get(0)->get_Result());
    
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormCheckBox, formFields->idx_get(1)->get_Type());
    ASSERT_TRUE(formFields->idx_get(1)->get_IsCheckBoxExactSize());
    ASSERT_EQ(u"Right click to check this box", formFields->idx_get(1)->get_HelpText());
    ASSERT_TRUE(formFields->idx_get(1)->get_OwnHelp());
    ASSERT_EQ(u"Checkbox status text", formFields->idx_get(1)->get_StatusText());
    ASSERT_TRUE(formFields->idx_get(1)->get_OwnStatus());
    ASPOSE_ASSERT_EQ(50.0, formFields->idx_get(1)->get_CheckBoxSize());
    ASSERT_FALSE(formFields->idx_get(1)->get_Checked());
    ASSERT_FALSE(formFields->idx_get(1)->get_Default());
    ASSERT_EQ(u"0", formFields->idx_get(1)->get_Result());
    
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormTextInput, formFields->idx_get(2)->get_Type());
    ASSERT_EQ(u"EntryMacro", formFields->idx_get(2)->get_EntryMacro());
    ASSERT_EQ(u"ExitMacro", formFields->idx_get(2)->get_ExitMacro());
    ASSERT_EQ(u"Regular", formFields->idx_get(2)->get_TextInputDefault());
    ASSERT_EQ(u"FIRST CAPITAL", formFields->idx_get(2)->get_TextInputFormat());
    ASSERT_EQ(Aspose::Words::Fields::TextFormFieldType::Regular, formFields->idx_get(2)->get_TextInputType());
    ASSERT_EQ(50, formFields->idx_get(2)->get_MaxLength());
    ASSERT_EQ(u"Regular", formFields->idx_get(2)->get_Result());
}


namespace gtest_test
{

class ExFormFields : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExFormFields> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExFormFields>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExFormFields> ExFormFields::s_instance;

} // namespace gtest_test

void ExFormFields::Create()
{
    //ExStart
    //ExFor:FormField
    //ExFor:FormField.Result
    //ExFor:FormField.Type
    //ExFor:FormField.Name
    //ExSummary:Shows how to insert a combo box.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Please select a fruit: ");
    
    // Insert a combo box which will allow a user to choose an option from a collection of strings.
    System::SharedPtr<Aspose::Words::Fields::FormField> comboBox = builder->InsertComboBox(u"MyComboBox", System::MakeArray<System::String>({u"Apple", u"Banana", u"Cherry"}), 0);
    
    ASSERT_EQ(u"MyComboBox", comboBox->get_Name());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormDropDown, comboBox->get_Type());
    ASSERT_EQ(u"Apple", comboBox->get_Result());
    
    // The form field will appear in the form of a "select" html tag.
    doc->Save(get_ArtifactsDir() + u"FormFields.Create.html");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"FormFields.Create.html");
    comboBox = doc->get_Range()->get_FormFields()->idx_get(0);
    
    ASSERT_EQ(u"MyComboBox", comboBox->get_Name());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldFormDropDown, comboBox->get_Type());
    ASSERT_EQ(u"Apple", comboBox->get_Result());
}

namespace gtest_test
{

TEST_F(ExFormFields, Create)
{
    s_instance->Create();
}

} // namespace gtest_test

void ExFormFields::TextInput()
{
    //ExStart
    //ExFor:DocumentBuilder.InsertTextInput
    //ExSummary:Shows how to insert a text input form field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Please enter text here: ");
    
    // Insert a text input field, which will allow the user to click it and enter text.
    // Assign some placeholder text that the user may overwrite and pass
    // a maximum text length of 0 to apply no limit on the form field's contents.
    builder->InsertTextInput(u"TextInput1", Aspose::Words::Fields::TextFormFieldType::Regular, u"", u"Placeholder text", 0);
    
    // The form field will appear in the form of an "input" html tag, with a type of "text".
    doc->Save(get_ArtifactsDir() + u"FormFields.TextInput.html");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"FormFields.TextInput.html");
    
    System::SharedPtr<Aspose::Words::Fields::FormField> textInput = doc->get_Range()->get_FormFields()->idx_get(0);
    
    ASSERT_EQ(u"TextInput1", textInput->get_Name());
    ASSERT_EQ(Aspose::Words::Fields::TextFormFieldType::Regular, textInput->get_TextInputType());
    ASSERT_EQ(System::String::Empty, textInput->get_TextInputFormat());
    ASSERT_EQ(u"Placeholder text", textInput->get_Result());
    ASSERT_EQ(0, textInput->get_MaxLength());
}

namespace gtest_test
{

TEST_F(ExFormFields, TextInput)
{
    s_instance->TextInput();
}

} // namespace gtest_test

void ExFormFields::DeleteFormField()
{
    //ExStart
    //ExFor:FormField.RemoveField
    //ExSummary:Shows how to delete a form field.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Form fields.docx");
    
    System::SharedPtr<Aspose::Words::Fields::FormField> formField = doc->get_Range()->get_FormFields()->idx_get(3);
    formField->RemoveField();
    //ExEnd
    
    System::SharedPtr<Aspose::Words::Fields::FormField> formFieldAfter = doc->get_Range()->get_FormFields()->idx_get(3);
    
    ASSERT_TRUE(System::TestTools::IsNull(formFieldAfter));
}

namespace gtest_test
{

TEST_F(ExFormFields, DeleteFormField)
{
    s_instance->DeleteFormField();
}

} // namespace gtest_test

void ExFormFields::DeleteFormFieldAssociatedWithBookmark()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->StartBookmark(u"MyBookmark");
    builder->InsertTextInput(u"TextInput1", Aspose::Words::Fields::TextFormFieldType::Regular, u"TestFormField", u"SomeText", 0);
    builder->EndBookmark(u"MyBookmark");
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    
    System::SharedPtr<Aspose::Words::BookmarkCollection> bookmarkBeforeDeleteFormField = doc->get_Range()->get_Bookmarks();
    ASSERT_EQ(u"MyBookmark", bookmarkBeforeDeleteFormField->idx_get(0)->get_Name());
    
    System::SharedPtr<Aspose::Words::Fields::FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);
    formField->RemoveField();
    
    System::SharedPtr<Aspose::Words::BookmarkCollection> bookmarkAfterDeleteFormField = doc->get_Range()->get_Bookmarks();
    ASSERT_EQ(u"MyBookmark", bookmarkAfterDeleteFormField->idx_get(0)->get_Name());
}

namespace gtest_test
{

TEST_F(ExFormFields, DeleteFormFieldAssociatedWithBookmark)
{
    s_instance->DeleteFormFieldAssociatedWithBookmark();
}

} // namespace gtest_test

void ExFormFields::FormFieldFontFormatting()
{
    //ExStart
    //ExFor:FormField
    //ExSummary:Shows how to formatting the entire FormField, including the field value.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Form fields.docx");
    
    System::SharedPtr<Aspose::Words::Fields::FormField> formField = doc->get_Range()->get_FormFields()->idx_get(0);
    formField->get_Font()->set_Bold(true);
    formField->get_Font()->set_Size(24);
    formField->get_Font()->set_Color(System::Drawing::Color::get_Red());
    
    formField->set_Result(u"Aspose.FormField");
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    
    System::SharedPtr<Aspose::Words::Run> formFieldRun = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(1);
    
    ASSERT_EQ(u"Aspose.FormField", formFieldRun->get_Text());
    ASPOSE_ASSERT_EQ(true, formFieldRun->get_Font()->get_Bold());
    ASPOSE_ASSERT_EQ(24, formFieldRun->get_Font()->get_Size());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), formFieldRun->get_Font()->get_Color().ToArgb());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFormFields, FormFieldFontFormatting)
{
    s_instance->FormFieldFontFormatting();
}

} // namespace gtest_test

void ExFormFields::Visitor()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Use a document builder to insert a combo box.
    builder->Write(u"Choose a value from this combo box: ");
    System::SharedPtr<Aspose::Words::Fields::FormField> comboBox = builder->InsertComboBox(u"MyComboBox", System::MakeArray<System::String>({u"One", u"Two", u"Three"}), 0);
    comboBox->set_CalculateOnExit(true);
    ASSERT_EQ(3, comboBox->get_DropDownItems()->get_Count());
    ASSERT_EQ(0, comboBox->get_DropDownSelectedIndex());
    ASSERT_TRUE(comboBox->get_Enabled());
    
    builder->InsertBreak(Aspose::Words::BreakType::ParagraphBreak);
    
    // Use a document builder to insert a check box.
    builder->Write(u"Click this check box to tick/untick it: ");
    System::SharedPtr<Aspose::Words::Fields::FormField> checkBox = builder->InsertCheckBox(u"MyCheckBox", false, 50);
    checkBox->set_IsCheckBoxExactSize(true);
    checkBox->set_HelpText(u"Right click to check this box");
    checkBox->set_OwnHelp(true);
    checkBox->set_StatusText(u"Checkbox status text");
    checkBox->set_OwnStatus(true);
    ASPOSE_ASSERT_EQ(50.0, checkBox->get_CheckBoxSize());
    ASSERT_FALSE(checkBox->get_Checked());
    ASSERT_FALSE(checkBox->get_Default());
    
    builder->InsertBreak(Aspose::Words::BreakType::ParagraphBreak);
    
    // Use a document builder to insert text input form field.
    builder->Write(u"Enter text here: ");
    System::SharedPtr<Aspose::Words::Fields::FormField> textInput = builder->InsertTextInput(u"MyTextInput", Aspose::Words::Fields::TextFormFieldType::Regular, u"", u"Placeholder text", 50);
    textInput->set_EntryMacro(u"EntryMacro");
    textInput->set_ExitMacro(u"ExitMacro");
    textInput->set_TextInputDefault(u"Regular");
    textInput->set_TextInputFormat(u"FIRST CAPITAL");
    textInput->SetTextInputValue(System::ExplicitCast<System::Object>(u"New placeholder text"));
    ASSERT_EQ(Aspose::Words::Fields::TextFormFieldType::Regular, textInput->get_TextInputType());
    ASSERT_EQ(50, textInput->get_MaxLength());
    
    // This collection contains all our form fields.
    System::SharedPtr<Aspose::Words::Fields::FormFieldCollection> formFields = doc->get_Range()->get_FormFields();
    ASSERT_EQ(3, formFields->get_Count());
    
    // Fields display our form fields. We can see their field codes by opening this document
    // in Microsoft and pressing Alt + F9. These fields have no switches,
    // and members of the FormField object fully govern their form fields' content.
    ASSERT_EQ(3, doc->get_Range()->get_Fields()->get_Count());
    ASSERT_EQ(u" FORMDROPDOWN \u0001", doc->get_Range()->get_Fields()->idx_get(0)->GetFieldCode());
    ASSERT_EQ(u" FORMCHECKBOX \u0001", doc->get_Range()->get_Fields()->idx_get(1)->GetFieldCode());
    ASSERT_EQ(u" FORMTEXT \u0001", doc->get_Range()->get_Fields()->idx_get(2)->GetFieldCode());
    
    // Allow each form field to accept a document visitor.
    auto formFieldVisitor = System::MakeObject<Aspose::Words::ApiExamples::ExFormFields::FormFieldVisitor>();
    
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Fields::FormField>>> fieldEnumerator = formFields->GetEnumerator();
        while (fieldEnumerator->MoveNext())
        {
            fieldEnumerator->get_Current()->Accept(formFieldVisitor);
        }
    }
    
    std::cout << formFieldVisitor->GetText() << std::endl;
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"FormFields.Visitor.html");
    TestFormField(doc);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExFormFields, Visitor)
{
    s_instance->Visitor();
}

} // namespace gtest_test

void ExFormFields::DropDownItemCollection()
{
    //ExStart
    //ExFor:DropDownItemCollection
    //ExFor:DropDownItemCollection.Add(String)
    //ExFor:DropDownItemCollection.Clear
    //ExFor:DropDownItemCollection.Contains(String)
    //ExFor:DropDownItemCollection.Count
    //ExFor:DropDownItemCollection.GetEnumerator
    //ExFor:DropDownItemCollection.IndexOf(String)
    //ExFor:DropDownItemCollection.Insert(Int32, String)
    //ExFor:DropDownItemCollection.Item(Int32)
    //ExFor:DropDownItemCollection.Remove(String)
    //ExFor:DropDownItemCollection.RemoveAt(Int32)
    //ExSummary:Shows how to insert a combo box field, and edit the elements in its item collection.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a combo box, and then verify its collection of drop-down items.
    // In Microsoft Word, the user will click the combo box,
    // and then choose one of the items of text in the collection to display.
    System::ArrayPtr<System::String> items = System::MakeArray<System::String>({u"One", u"Two", u"Three"});
    System::SharedPtr<Aspose::Words::Fields::FormField> comboBoxField = builder->InsertComboBox(u"DropDown", items, 0);
    System::SharedPtr<Aspose::Words::Fields::DropDownItemCollection> dropDownItems = comboBoxField->get_DropDownItems();
    
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
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::String>> dropDownCollectionEnumerator = dropDownItems->GetEnumerator();
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
    
    doc->Save(get_ArtifactsDir() + u"FormFields.DropDownItemCollection.html");
    
    // Empty the whole collection of drop-down items.
    dropDownItems->Clear();
    //ExEnd
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    dropDownItems = doc->get_Range()->get_FormFields()->idx_get(0)->get_DropDownItems();
    
    ASSERT_EQ(0, dropDownItems->get_Count());
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"FormFields.DropDownItemCollection.html");
    dropDownItems = doc->get_Range()->get_FormFields()->idx_get(0)->get_DropDownItems();
    
    ASSERT_EQ(3, dropDownItems->get_Count());
    ASSERT_EQ(u"One", dropDownItems->idx_get(0));
    ASSERT_EQ(u"Two", dropDownItems->idx_get(1));
    ASSERT_EQ(u"Three", dropDownItems->idx_get(2));
}

namespace gtest_test
{

TEST_F(ExFormFields, DropDownItemCollection)
{
    s_instance->DropDownItemCollection();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
