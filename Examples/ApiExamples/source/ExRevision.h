#pragma once

#include <system/string.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionType.h>
#include <Aspose.Words.Cpp/Model/Revisions/Revision.h>
#include <Aspose.Words.Cpp/Model/Revisions/IRevisionCriteria.h>
#include <Aspose.Words.Cpp/Model/Comparing/Granularity.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Comparing;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExRevision : public ApiExampleBase
{
    typedef ExRevision ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Control when certain revision should be accepted/rejected.
    /// </summary>
    class RevisionCriteria : public IRevisionCriteria
    {
        typedef RevisionCriteria ThisType;
        typedef IRevisionCriteria BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        RevisionCriteria(System::String authorName, Aspose::Words::RevisionType revisionType);
        
        bool IsMatch(System::SharedPtr<Aspose::Words::Revision> revision) override;
        
    private:
    
        System::String AuthorName;
        Aspose::Words::RevisionType RevisionType;
        
    };
    
    
public:

    void Revisions();
    void RevisionCollection();
    void GetInfoAboutRevisionsInRevisionGroups();
    void GetSpecificRevisionGroup();
    void ShowRevisionBalloons();
    void RevisionOptions();
    //ExStart:RevisionSpecifiedCriteria
    //GistId:470c0da51e4317baae82ad9495747fed
    //ExFor:RevisionCollection.Accept(IRevisionCriteria)
    //ExFor:RevisionCollection.Reject(IRevisionCriteria)
    //ExFor:IRevisionCriteria
    //ExFor:IRevisionCriteria.IsMatch(Revision)
    //ExSummary:Shows how to accept or reject revision based on criteria.
    void RevisionSpecifiedCriteria();
    //ExEnd:RevisionSpecifiedCriteria
    void TrackRevisions();
    void AcceptAllRevisions();
    void GetRevisedPropertiesOfList();
    void Compare();
    void CompareDocumentWithRevisions();
    void CompareOptions();
    void IgnoreDmlUniqueId(bool isIgnoreDmlUniqueId);
    void LayoutOptionsRevisions();
    void GranularityCompareOption(Aspose::Words::Comparing::Granularity granularity);
    void IgnoreStoreItemId();
    void RevisionCellColor();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


