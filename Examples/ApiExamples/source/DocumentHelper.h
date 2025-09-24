#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/array.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Tables;

namespace Aspose {

namespace Words {

namespace ApiExamples {

/// <summary>
/// Functions for operations with documents and content
/// </summary>
class DocumentHelper : public ApiExampleBase
{
    typedef DocumentHelper ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Create simple document without run in the paragraph
    /// </summary>
    static System::SharedPtr<Aspose::Words::Document> CreateDocumentWithoutDummyText();
    /// <summary>
    /// Create new document with text
    /// </summary>
    static System::SharedPtr<Aspose::Words::Document> CreateDocumentFillWithDummyText();
    static void FindTextInFile(System::String path, System::String expression);
    /// <summary>
    /// Create new document template for reporting engine
    /// </summary>
    static System::SharedPtr<Aspose::Words::Document> CreateSimpleDocument(System::String templateText);
    /// <summary>
    /// Create new document with textbox shape and some query
    /// </summary>
    static System::SharedPtr<Aspose::Words::Document> CreateTemplateDocumentWithDrawObjects(System::String templateText, Aspose::Words::Drawing::ShapeType shapeType);
    /// <summary>
    /// Compare word documents
    /// </summary>
    /// <param name="filePathDoc1">First document path</param>
    /// <param name="filePathDoc2">Second document path</param>
    /// <returns>Result of compare document</returns>
    static bool CompareDocs(System::String filePathDoc1, System::String filePathDoc2);
    /// <summary>
    /// Insert run into the current document
    /// </summary>
    /// <param name="doc">Current document</param>
    /// <param name="text">Custom text</param>
    /// <param name="paraIndex">Paragraph index</param>
    static System::SharedPtr<Aspose::Words::Run> InsertNewRun(System::SharedPtr<Aspose::Words::Document> doc, System::String text, int32_t paraIndex);
    /// <summary>
    /// Insert text into the current document
    /// </summary>
    /// <param name="builder">Current document builder</param>
    /// <param name="textStrings">Custom text</param>
    static void InsertBuilderText(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::ArrayPtr<System::String> textStrings);
    /// <summary>
    /// Get paragraph text of the current document
    /// </summary>
    /// <param name="doc">Current document</param>
    /// <param name="paraIndex">Paragraph number from collection</param>
    static System::String GetParagraphText(System::SharedPtr<Aspose::Words::Document> doc, int32_t paraIndex);
    /// <summary>
    /// Insert new table in the document
    /// </summary>
    /// <param name="builder">Current document builder</param>
    static System::SharedPtr<Aspose::Words::Tables::Table> InsertTable(System::SharedPtr<Aspose::Words::DocumentBuilder> builder);
    /// <summary>
    /// Insert TOC entries in the document
    /// </summary>
    /// <param name="builder">
    /// The builder.
    /// </param>
    static void InsertToc(System::SharedPtr<Aspose::Words::DocumentBuilder> builder);
    /// <summary>
    /// Get section text of the current document
    /// </summary>
    /// <param name="doc">Current document</param>
    /// <param name="secIndex">Section number from collection</param>
    static System::String GetSectionText(System::SharedPtr<Aspose::Words::Document> doc, int32_t secIndex);
    /// <summary>
    /// Get paragraph of the current document
    /// </summary>
    /// <param name="doc">Current document</param>
    /// <param name="paraIndex">Paragraph number from collection</param>
    static System::SharedPtr<Aspose::Words::Paragraph> GetParagraph(System::SharedPtr<Aspose::Words::Document> doc, int32_t paraIndex);
    /// <summary>
    /// Save the document to a stream, immediately re-open it and return the newly opened version
    /// </summary>
    /// <remarks>
    /// Used for testing how document features are preserved after saving/loading
    /// </remarks>
    /// <param name="doc">The document we wish to re-open</param>
    static System::SharedPtr<Aspose::Words::Document> SaveOpen(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


