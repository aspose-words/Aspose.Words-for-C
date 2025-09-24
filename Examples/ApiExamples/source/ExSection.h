#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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


