#pragma once

#include <system/shared_ptr.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Markup/SmartTag.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Markup;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExSmartTag : public ApiExampleBase
{
    typedef ExSmartTag ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Prints visited smart tags and their contents.
    /// </summary>
    class SmartTagPrinter : public DocumentVisitor
    {
        typedef SmartTagPrinter ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// Called when a SmartTag node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitSmartTagStart(System::SharedPtr<Aspose::Words::Markup::SmartTag> smartTag) override;
        /// <summary>
        /// Called when the visiting of a SmartTag node is ended.
        /// </summary>
        Aspose::Words::VisitorAction VisitSmartTagEnd(System::SharedPtr<Aspose::Words::Markup::SmartTag> smartTag) override;
        
    };
    
    
public:

    //ExStart
    //ExFor:CompositeNode.RemoveSmartTags
    //ExFor:CustomXmlProperty
    //ExFor:CustomXmlProperty.#ctor(String,String,String)
    //ExFor:CustomXmlProperty.Name
    //ExFor:CustomXmlProperty.Value
    //ExFor:SmartTag
    //ExFor:SmartTag.#ctor(DocumentBase)
    //ExFor:SmartTag.Accept(DocumentVisitor)
    //ExFor:SmartTag.AcceptStart(DocumentVisitor)
    //ExFor:SmartTag.AcceptEnd(DocumentVisitor)
    //ExFor:SmartTag.Element
    //ExFor:SmartTag.Properties
    //ExFor:SmartTag.Uri
    //ExSummary:Shows how to create smart tags.
    void Create();
    //ExEnd
    void TestCreate(System::SharedPtr<Aspose::Words::Document> doc);
    void Properties();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


