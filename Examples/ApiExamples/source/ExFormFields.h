#pragma once

#include <system/text/string_builder.h>
#include <system/string.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Fields;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExFormFields : public ApiExampleBase
{
    typedef ExFormFields ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Visitor implementation that prints details of form fields that it visits.
    /// </summary>
    class FormFieldVisitor : public DocumentVisitor
    {
        typedef FormFieldVisitor ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        FormFieldVisitor();
        
        /// <summary>
        /// Called when a FormField node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFormField(System::SharedPtr<Aspose::Words::Fields::FormField> formField) override;
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Adds newline char-terminated text to the current output.
        /// </summary>
        void AppendLine(System::String text);
        
    };
    
    
public:

    void Create();
    void TextInput();
    void DeleteFormField();
    void DeleteFormFieldAssociatedWithBookmark();
    void FormFieldFontFormatting();
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
    void Visitor();
    void DropDownItemCollection();
    
protected:

    //ExEnd
    void TestFormField(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


