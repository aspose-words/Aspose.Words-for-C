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
namespace Aspose { namespace Words { class EditableRangeStart; } }
namespace Aspose { namespace Words { class EditableRangeEnd; } }
namespace Aspose { namespace Words { class Run; } }
namespace Aspose { namespace Words { class Document; } }

namespace ApiExamples {

class ExEditableRange : public ApiExampleBase
{
public:

    /// <summary>
    /// Visitor implementation that prints attributes and contents of ranges.
    /// </summary>
    class EditableRangeInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef EditableRangeInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        EditableRangeInfoPrinter();
        
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
        /// Called when a Run node is encountered in the document. Only runs within editable ranges have their contents recorded.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        bool mInsideEditableRange;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
    };
    
    
public:

    void RemovesEditableRange();
    void CreateEditableRanges();
    void IncorrectStructureException();
    void IncorrectStructureDoNotAdded();
    
protected:

    void TestCreateEditableRanges(System::SharedPtr<Aspose::Words::Document> doc, System::SharedPtr<ExEditableRange::EditableRangeInfoPrinter> visitor);
    
};

} // namespace ApiExamples


