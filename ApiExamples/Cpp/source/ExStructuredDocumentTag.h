#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/test_tools/method_argument_tuple.h>

#include "ApiExampleBase.h"

namespace ApiExamples {

/// <summary>
/// Tests that verify work with structured document tags in the document. 
/// </summary>
class ExStructuredDocumentTag : public ApiExampleBase
{
public:

    void RepeatingSection();
    void SetSpecificStyleToSdt();
    void CheckBox();
    void Date();
    void PlainText();
    void IsTemporary();
    void PlaceholderBuildingBlock();
    void Lock();
    void ListItemCollection();
    void CreatingCustomXml();
    void XmlMapping();
    void CustomXmlSchemaCollection();
    void CustomXmlPartStoreItemIdReadOnly();
    void CustomXmlPartStoreItemIdReadOnlyNull();
    void ClearTextFromStructuredDocumentTags();
    void AccessToBuildingBlockPropertiesFromDocPartObjSdt();
    void AccessToBuildingBlockPropertiesFromPlainTextSdt();
    void BuildingBlockCategories();
    void UpdateSdtContent(bool updateSdtContent);
    void FillTableUsingRepeatingSectionItem();
    void CustomXmlPart();
    
};

} // namespace ApiExamples


