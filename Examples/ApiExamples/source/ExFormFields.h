#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
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
    void Visitor();
    void DropDownItemCollection();
    
protected:

    void TestFormField(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


