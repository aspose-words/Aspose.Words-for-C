#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <system/guid.h>
#include <system/collections/dictionary.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/GlossaryDocument.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlock.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::BuildingBlocks;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExBuildingBlocks : public ApiExampleBase
{
    typedef ExBuildingBlocks ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Sets up a visited building block to be inserted into the document as a quick part and adds text to its contents.
    /// </summary>
    class BuildingBlockVisitor : public DocumentVisitor
    {
        typedef BuildingBlockVisitor ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        BuildingBlockVisitor(System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> ownerGlossaryDoc);
        
        Aspose::Words::VisitorAction VisitBuildingBlockStart(System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block) override;
        Aspose::Words::VisitorAction VisitBuildingBlockEnd(System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block) override;
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> mGlossaryDoc;
        
    };
    
    /// <summary>
    /// Gives each building block in a visited glossary document a unique GUID.
    /// Stores the GUID-building block pairs in a dictionary.
    /// </summary>
    class GlossaryDocVisitor : public DocumentVisitor
    {
        typedef GlossaryDocVisitor ThisType;
        typedef DocumentVisitor BaseType;
        
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
        
    private:
    
        System::SharedPtr<System::Collections::Generic::Dictionary<System::Guid, System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>> mBlocksByGuid;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
    };
    
    
public:

    void CreateAndInsert();
    void GlossaryDocument();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


