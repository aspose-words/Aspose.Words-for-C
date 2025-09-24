#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/shared_ptr.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTagRangeStart.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Markup;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExStructuredDocumentTag : public ApiExampleBase
{
    typedef ExStructuredDocumentTag ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void RepeatingSection();
    void FlatOpcContent();
    void ApplyStyle();
    void CheckBox();
    void Date();
    void PlainText();
    void IsTemporary(bool isTemporary);
    void PlaceholderBuildingBlock(bool isShowingPlaceholderText);
    void Lock();
    void ListItemCollection();
    void CreatingCustomXml();
    void DataChecksum();
    void XmlMapping();
    void StructuredDocumentTagRangeStartXmlMapping();
    void CustomXmlSchemaCollection();
    void CustomXmlPartStoreItemIdReadOnly();
    void CustomXmlPartStoreItemIdReadOnlyNull();
    void ClearTextFromStructuredDocumentTags();
    void AccessToBuildingBlockPropertiesFromDocPartObjSdt();
    void AccessToBuildingBlockPropertiesFromPlainTextSdt();
    void BuildingBlockCategories();
    void UpdateSdtContent();
    void FillTableUsingRepeatingSectionItem();
    void CustomXmlPart();
    void MultiSectionTags();
    void SdtChildNodes();
    void SdtRangeExtendedMethods();
    System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTagRangeStart> InsertStructuredDocumentTagRanges(System::SharedPtr<Aspose::Words::Document> doc);
    void GetSdt();
    void RangeSdt();
    void SdtAtRowLevel();
    void IgnoreStructuredDocumentTags();
    void Citation();
    void RangeStartWordOpenXmlMinimal();
    void RemoveSelfOnly();
    void Appearance();
    void InsertStructuredDocumentTag();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


