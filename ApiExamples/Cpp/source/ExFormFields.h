#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { enum class VisitorAction; } }
namespace Aspose { namespace Words { namespace Fields { class FormField; } } }
namespace Aspose { namespace Words { class Document; } }

namespace ApiExamples {

class ExFormFields : public ApiExampleBase
{
public:

    /// <summary>
    /// Visitor implementation that prints information about visited form fields. 
    /// </summary>
    class FormFieldVisitor : public Aspose::Words::DocumentVisitor
    {
        typedef FormFieldVisitor ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
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
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Adds newline char-terminated text to the current output.
        /// </summary>
        void AppendLine(System::String text);
        
    };
    
    
public:

    void FormFieldsWorkWithProperties();
    void InsertAndRetrieveFormFields();
    void DeleteFormField();
    void DeleteFormFieldAssociatedWithBookmark();
    void FormField();
    
protected:

    void TestFormField(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples


