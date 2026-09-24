#pragma once

#include <system/text/string_builder.h>
#include <system/string.h>
#include <gtest/gtest.h>
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
    //ExStart
    //ExFor:EditableRange
    //ExFor:EditableRange.EditorGroup
    //ExFor:EditableRange.SingleUser
    //ExFor:EditableRangeEnd
    //ExFor:EditableRangeEnd.Accept(DocumentVisitor)
    //ExFor:EditableRangeStart
    //ExFor:EditableRangeStart.Accept(DocumentVisitor)
    //ExFor:EditorType
    //ExSummary:Shows how to limit the editing rights of editable ranges to a specific group/user.
    void Visitor();
    //ExEnd
    void IncorrectStructureException();
    void IncorrectStructureDoNotAdded();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


