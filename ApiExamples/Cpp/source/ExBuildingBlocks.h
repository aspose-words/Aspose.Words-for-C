#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <system/guid.h>
#include <system/collections/dictionary.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace BuildingBlocks { class GlossaryDocument; } } }
namespace Aspose { namespace Words { enum class VisitorAction; } }
namespace Aspose { namespace Words { namespace BuildingBlocks { class BuildingBlock; } } }

namespace ApiExamples {

class ExBuildingBlocks : public ApiExampleBase
{
public:

    /// <summary>
    /// Simple implementation of adding text to a building block and preparing it for usage in the document. Implemented as a Visitor.
    /// </summary>
    class BuildingBlockVisitor : public Aspose::Words::DocumentVisitor
    {
        typedef BuildingBlockVisitor ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        BuildingBlockVisitor(System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> ownerGlossaryDoc);
        
        Aspose::Words::VisitorAction VisitBuildingBlockStart(System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block) override;
        Aspose::Words::VisitorAction VisitBuildingBlockEnd(System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> mGlossaryDoc;
        
    };
    
    /// <summary>
    /// Simple implementation of giving each building block in a glossary document a unique GUID. Implemented as a Visitor.
    /// </summary>
    class GlossaryDocVisitor : public Aspose::Words::DocumentVisitor
    {
        typedef GlossaryDocVisitor ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        GlossaryDocVisitor();
        
        System::String GetText();
        System::SharedPtr<System::Collections::Generic::Dictionary<System::Guid, System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>> GetDictionary();
        Aspose::Words::VisitorAction VisitGlossaryDocumentStart(System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> glossary) override;
        Aspose::Words::VisitorAction VisitGlossaryDocumentEnd(System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> glossary) override;
        Aspose::Words::VisitorAction VisitBuildingBlockStart(System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block) override;
        Aspose::Words::VisitorAction VisitBuildingBlockEnd(System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Collections::Generic::Dictionary<System::Guid, System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>> mBlocksByGuid;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
    };
    
    
public:

    void BuildingBlockFields();
    void GlossaryDocument();
    
};

} // namespace ApiExamples


