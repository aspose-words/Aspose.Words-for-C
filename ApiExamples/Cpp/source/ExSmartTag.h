#pragma once

#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { enum class VisitorAction; } }
namespace Aspose { namespace Words { namespace Markup { class SmartTag; } } }
namespace Aspose { namespace Words { class Document; } }

namespace ApiExamples {

class ExSmartTag : public ApiExampleBase
{
private:

    /// <summary>
    /// DocumentVisitor implementation that prints smart tags and their contents
    /// </summary>
    class SmartTagVisitor : public Aspose::Words::DocumentVisitor
    {
        typedef SmartTagVisitor ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
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

    void Create();
    void TestCreate(System::SharedPtr<Aspose::Words::Document> doc);
    void Properties();
    
};

} // namespace ApiExamples


