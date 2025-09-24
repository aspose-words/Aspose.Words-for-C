#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Fields;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExParagraph : public ApiExampleBase
{
    typedef ExParagraph ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void DocumentBuilderInsertParagraph();
    void AppendField();
    void InsertField();
    void InsertFieldBeforeTextInParagraph();
    void InsertFieldAfterTextInParagraph();
    void InsertFieldBeforeTextInParagraphWithoutUpdateField();
    void InsertFieldAfterTextInParagraphWithoutUpdateField();
    void InsertFieldWithoutSeparator();
    void InsertFieldBeforeParagraphWithoutDocumentAuthor();
    void InsertFieldAfterParagraphWithoutChangingDocumentAuthor();
    void InsertFieldBeforeRunText();
    void InsertFieldAfterRunText();
    void InsertFieldEmptyParagraphWithoutUpdateField();
    void InsertFieldEmptyParagraphWithUpdateField();
    void CompositeNodeChildren();
    void MoveRevisions();
    void RangeRevisions();
    void GetFormatRevision();
    void GetFrameProperties();
    void IsRevision();
    void BreakIsStyleSeparator();
    void TabStops();
    void JoinRuns();
    
protected:

    /// <summary>
    /// Insert field into the first paragraph of the current document using field type.
    /// </summary>
    static void InsertFieldUsingFieldType(System::SharedPtr<Aspose::Words::Document> doc, Aspose::Words::Fields::FieldType fieldType, bool updateField, System::SharedPtr<Aspose::Words::Node> refNode, bool isAfter, int32_t paraIndex);
    /// <summary>
    /// Insert field into the first paragraph of the current document using field code.
    /// </summary>
    static void InsertFieldUsingFieldCode(System::SharedPtr<Aspose::Words::Document> doc, System::String fieldCode, System::SharedPtr<Aspose::Words::Node> refNode, bool isAfter, int32_t paraIndex);
    /// <summary>
    /// Insert field into the first paragraph of the current document using field code and field String.
    /// </summary>
    static void InsertFieldUsingFieldCodeFieldString(System::SharedPtr<Aspose::Words::Document> doc, System::String fieldCode, System::String fieldValue, System::SharedPtr<Aspose::Words::Node> refNode, bool isAfter, int32_t paraIndex);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


