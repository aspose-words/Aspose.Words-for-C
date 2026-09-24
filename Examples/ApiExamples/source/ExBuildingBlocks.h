#pragma once

#include <system/text/string_builder.h>
#include <system/string.h>
#include <system/guid.h>
#include <system/collections/dictionary.h>
#include <gtest/gtest.h>
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

    //ExStart
    //ExFor:Document.GlossaryDocument
    //ExFor:BuildingBlock
    //ExFor:BuildingBlock.#ctor(GlossaryDocument)
    //ExFor:BuildingBlock.Accept(DocumentVisitor)
    //ExFor:BuildingBlock.AcceptStart(DocumentVisitor)
    //ExFor:BuildingBlock.AcceptEnd(DocumentVisitor)
    //ExFor:BuildingBlock.Behavior
    //ExFor:BuildingBlock.Category
    //ExFor:BuildingBlock.Description
    //ExFor:BuildingBlock.FirstSection
    //ExFor:BuildingBlock.Gallery
    //ExFor:BuildingBlock.Guid
    //ExFor:BuildingBlock.LastSection
    //ExFor:BuildingBlock.Name
    //ExFor:BuildingBlock.Sections
    //ExFor:BuildingBlock.Type
    //ExFor:BuildingBlockBehavior
    //ExFor:BuildingBlockType
    //ExSummary:Shows how to add a custom building block to a document.
    void CreateAndInsert();
    //ExEnd
    //ExStart
    //ExFor:GlossaryDocument
    //ExFor:GlossaryDocument.Accept(DocumentVisitor)
    //ExFor:GlossaryDocument.AcceptStart(DocumentVisitor)
    //ExFor:GlossaryDocument.AcceptEnd(DocumentVisitor)
    //ExFor:GlossaryDocument.BuildingBlocks
    //ExFor:GlossaryDocument.FirstBuildingBlock
    //ExFor:GlossaryDocument.GetBuildingBlock(BuildingBlockGallery,String,String)
    //ExFor:GlossaryDocument.LastBuildingBlock
    //ExFor:BuildingBlockCollection
    //ExFor:BuildingBlockCollection.Item(Int32)
    //ExFor:BuildingBlockCollection.ToArray
    //ExFor:BuildingBlockGallery
    //ExFor:DocumentVisitor.VisitBuildingBlockEnd(BuildingBlock)
    //ExFor:DocumentVisitor.VisitBuildingBlockStart(BuildingBlock)
    //ExFor:DocumentVisitor.VisitGlossaryDocumentEnd(GlossaryDocument)
    //ExFor:DocumentVisitor.VisitGlossaryDocumentStart(GlossaryDocument)
    //ExSummary:Shows ways of accessing building blocks in a glossary document.
    void GlossaryDocument();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


