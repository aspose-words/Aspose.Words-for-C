#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeStart.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExEditableRange : public ApiExampleBase
{
    typedef ExEditableRange ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Collects properties and contents of visited editable ranges in a string.
    /// </summary>
    class EditableRangePrinter : public DocumentVisitor
    {
        typedef EditableRangePrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        EditableRangePrinter();
        
        System::String ToText();
        void Reset();
        /// <summary>
        /// Called when an EditableRangeStart node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitEditableRangeStart(System::SharedPtr<Aspose::Words::EditableRangeStart> editableRangeStart) override;
        /// <summary>
        /// Called when an EditableRangeEnd node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitEditableRangeEnd(System::SharedPtr<Aspose::Words::EditableRangeEnd> editableRangeEnd) override;
        /// <summary>
        /// Called when a Run node is encountered in the document. This visitor only records runs that are inside editable ranges.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        
    private:
    
        bool mInsideEditableRange;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
    };
    
    
public:

    void CreateAndRemove();
    void Nested();
    void Visitor();
    void IncorrectStructureException();
    void IncorrectStructureDoNotAdded();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


