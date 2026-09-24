#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExSection : public ApiExampleBase
{
    typedef ExSection ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void Protect();
    void AddRemove();
    void FirstAndLast();
    void CreateManually();
    void EnsureMinimum();
    void BodyEnsureMinimum();
    void BodyChildNodes();
    void Clear();
    void PrependAppendContent();
    void ClearContent();
    void ClearHeadersFooters();
    void DeleteHeaderFooterShapes();
    void SectionsCloneSection();
    void SectionsImportSection();
    void MigrateFrom2XImportSection();
    void ModifyPageSetupInAllSections();
    void CultureInfoPageSetupDefaults();
    void PreserveWatermarks();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


