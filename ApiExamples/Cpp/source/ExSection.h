#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include "ApiExampleBase.h"

namespace ApiExamples {

class ExSection : public ApiExampleBase
{
public:

    void Protect();
    void AddRemove();
    void CreateFromScratch();
    void EnsureSectionMinimum();
    void BodyEnsureMinimum();
    void BodyNodeType();
    void SectionsDeleteAllSections();
    void SectionsAppendSectionContent();
    void SectionsDeleteSectionContent();
    void SectionsDeleteHeaderFooter();
    void SectionDeleteHeaderFooterShapes();
    void SectionsCloneSection();
    void SectionsImportSection();
    void MigrateFrom2XImportSection();
    void ModifyPageSetupInAllSections();
    void CultureInfoPageSetupDefaults();
    
};

} // namespace ApiExamples


