#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/timespan.h>
#include <system/text/encoding.h>
#include <system/io/stream.h>
#include <system/date_time.h>
#include <system/collections/dictionary.h>
#include <system/array.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/TabStop.h>
#include <Aspose.Words.Cpp/Model/Text/TabLeader.h>
#include <Aspose.Words.Cpp/Model/Text/TabAlignment.h>
#include <Aspose.Words.Cpp/Model/Numbering/NumberStyle.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditorType.h>
#include <Aspose.Words.Cpp/Model/EditableRanges/EditableRange.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBoxWrapMode.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBox.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/LayoutFlow.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Notes;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class TestUtil : public ApiExampleBase
{
    typedef TestUtil ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Checks whether a file at a specified filename contains a valid image with specified dimensions.
    /// </summary>
    /// <remarks>
    /// Serves to check that an image file is valid and nonempty without looking up its file size.
    /// </remarks>
    /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
    /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
    /// <param name="filename">Local file system filename of the image file.</param>
    static void VerifyImage(int32_t expectedWidth, int32_t expectedHeight, System::String filename);
    static System::SharedPtr<System::Collections::Generic::Dictionary<System::String, int32_t>> GetEmfDimensions(System::String filePath);
    static System::SharedPtr<System::Collections::Generic::Dictionary<System::String, int32_t>> GetWmfDimensions(System::String filePath);
    /// <summary>
    /// Checks whether an image from the local file system contains any transparency.
    /// </summary>
    /// <param name="filename">Local file system filename of the image file.</param>
    static void ImageContainsTransparency(System::String filename);
    /// <summary>
    /// Checks whether a document produced during a mail merge contains every element of an array of arrays of strings.
    /// </summary>
    /// <remarks>
    /// Only suitable for rectangular arrays.
    /// </remarks>
    /// <param name="expectedResult">Values from the mail merge data source that we expect the document to contain.</param>
    /// <param name="doc">Document created during a mail merge.</param>
    /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
    static void MailMergeMatchesArray(System::ArrayPtr<System::ArrayPtr<System::String>> expectedResult, System::SharedPtr<Aspose::Words::Document> doc, bool onePagePerRow);
    /// <summary>
    /// Checks whether a file in the local file system contains a string in its raw data.
    /// </summary>
    /// <param name="expected">The string we are looking for.</param>
    /// <param name="filename">Local system filename of a file which, when read from the beginning, should contain the string.</param>
    static void FileContainsString(System::String expected, System::String filename);
    /// <summary>
    /// Checks whether values of properties of a field with a type not related to date/time are equal to expected values.
    /// </summary>
    /// <remarks>
    /// Best used when there are many fields closely being tested and should be avoided if a field has a long field code/result.
    /// </remarks>
    /// <param name="expectedType">The FieldType that we expect the field to have.</param>
    /// <param name="expectedFieldCode">The expected output value of GetFieldCode() being called on the field.</param>
    /// <param name="expectedResult">The field's expected result, which will be the value displayed by it in the document.</param>
    /// <param name="field">The field that's being tested.</param>
    static void VerifyField(Aspose::Words::Fields::FieldType expectedType, System::String expectedFieldCode, System::String expectedResult, System::SharedPtr<Aspose::Words::Fields::Field> field);
    /// <summary>
    /// Checks whether values of properties of a field with a type related to date/time are equal to expected values.
    /// </summary>
    /// <remarks>
    /// Used when comparing DateTime instances to Field.Result values parsed to DateTime, which may differ slightly. 
    /// Give a delta value that's generous enough for any lower end system to pass, also a delta of zero is allowed.
    /// </remarks>
    /// <param name="expectedType">The FieldType that we expect the field to have.</param>
    /// <param name="expectedFieldCode">The expected output value of GetFieldCode() being called on the field.</param>
    /// <param name="expectedResult">The date/time that the field's result is expected to represent.</param>
    /// <param name="field">The field that's being tested.</param>
    /// <param name="delta">Margin of error for expectedResult.</param>
    static void VerifyField(Aspose::Words::Fields::FieldType expectedType, System::String expectedFieldCode, System::DateTime expectedResult, System::SharedPtr<Aspose::Words::Fields::Field> field, System::TimeSpan delta);
    /// <summary>
    /// Checks whether a DateTime matches an expected value, with a margin of error.
    /// </summary>
    /// <param name="expected">The date/time that we expect the result to be.</param>
    /// <param name="actual">The DateTime object being tested.</param>
    /// <param name="delta">Margin of error for expectedResult.</param>
    static void VerifyDate(System::DateTime expected, System::DateTime actual, System::TimeSpan delta);
    /// <summary>
    /// Checks whether a field contains another complete field as a sibling within its nodes.
    /// </summary>
    /// <remarks>
    /// If two fields have the same immediate parent node and therefore their nodes are siblings,
    /// the FieldStart of the outer field appears before the FieldStart of the inner node,
    /// and the FieldEnd of the outer node appears after the FieldEnd of the inner node,
    /// then the inner field is considered to be nested within the outer field. 
    /// </remarks>
    /// <param name="innerField">The field that we expect to be fully within outerField.</param>
    /// <param name="outerField">The field that we to contain innerField.</param>
    static void FieldsAreNested(System::SharedPtr<Aspose::Words::Fields::Field> innerField, System::SharedPtr<Aspose::Words::Fields::Field> outerField);
    /// <summary>
    /// Checks whether a shape contains a valid image with specified dimensions.
    /// </summary>
    /// <remarks>
    /// Serves to check that an image file is valid and nonempty without looking up its data length.
    /// </remarks>
    /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
    /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
    /// <param name="expectedImageType">Expected format of the image.</param>
    /// <param name="imageShape">Shape that contains the image.</param>
    static void VerifyImageInShape(int32_t expectedWidth, int32_t expectedHeight, Aspose::Words::Drawing::ImageType expectedImageType, System::SharedPtr<Aspose::Words::Drawing::Shape> imageShape);
    /// <summary>
    /// Checks whether values of a footnote's properties are equal to their expected values.
    /// </summary>
    /// <param name="expectedFootnoteType">Expected type of the footnote/endnote.</param>
    /// <param name="expectedIsAuto">Expected auto-numbered status of this footnote.</param>
    /// <param name="expectedReferenceMark">If "IsAuto" is false, then the footnote is expected to display this string instead of a number after referenced text.</param>
    /// <param name="expectedContents">Expected side comment provided by the footnote.</param>
    /// <param name="footnote">Footnote node in question.</param>
    static void VerifyFootnote(Aspose::Words::Notes::FootnoteType expectedFootnoteType, bool expectedIsAuto, System::String expectedReferenceMark, System::String expectedContents, System::SharedPtr<Aspose::Words::Notes::Footnote> footnote);
    /// <summary>
    /// Checks whether values of a list level's properties are equal to their expected values.
    /// </summary>
    /// <remarks>
    /// Only necessary for list levels that have been explicitly created by the user.
    /// </remarks>
    /// <param name="expectedListFormat">Expected format for the list symbol.</param>
    /// <param name="expectedNumberPosition">Expected indent for this level, usually growing larger with each level.</param>
    /// <param name="expectedNumberStyle"></param>
    /// <param name="listLevel">List level in question.</param>
    static void VerifyListLevel(System::String expectedListFormat, double expectedNumberPosition, Aspose::Words::NumberStyle expectedNumberStyle, System::SharedPtr<Aspose::Words::Lists::ListLevel> listLevel);
    /// <summary>
    /// Copies from the current position in src stream till the end.
    /// Copies into the current position in dst stream.
    /// </summary>
    static void CopyStream(System::SharedPtr<System::IO::Stream> srcStream, System::SharedPtr<System::IO::Stream> dstStream);
    /// <summary>
    /// Dumps byte array into a string.
    /// </summary>
    static System::String DumpArray(System::ArrayPtr<uint8_t> data, int32_t start, int32_t count);
    /// <summary>
    /// Checks whether values of a tab stop's properties are equal to their expected values.
    /// </summary>
    /// <param name="expectedPosition">Expected position on the tab stop ruler, in points.</param>
    /// <param name="expectedTabAlignment">Expected position where the position is measured from </param>
    /// <param name="expectedTabLeader">Expected characters that pad the space between the start and end of the tab whitespace.</param>
    /// <param name="isClear">Whether or no this tab stop clears any tab stops.</param>
    /// <param name="tabStop">Tab stop that's being tested.</param>
    static void VerifyTabStop(double expectedPosition, Aspose::Words::TabAlignment expectedTabAlignment, Aspose::Words::TabLeader expectedTabLeader, bool isClear, System::SharedPtr<Aspose::Words::TabStop> tabStop);
    /// <summary>
    /// Checks whether values of a shape's properties are equal to their expected values.
    /// </summary>
    /// <remarks>
    /// All dimension measurements are in points.
    /// </remarks>
    static void VerifyShape(Aspose::Words::Drawing::ShapeType expectedShapeType, System::String expectedName, double expectedWidth, double expectedHeight, double expectedTop, double expectedLeft, System::SharedPtr<Aspose::Words::Drawing::Shape> shape);
    /// <summary>
    /// Checks whether values of properties of a textbox are equal to their expected values.
    /// </summary>
    /// <remarks>
    /// All dimension measurements are in points.
    /// </remarks>
    static void VerifyTextBox(Aspose::Words::Drawing::LayoutFlow expectedLayoutFlow, bool expectedFitShapeToText, Aspose::Words::Drawing::TextBoxWrapMode expectedTextBoxWrapMode, double marginTop, double marginBottom, double marginLeft, double marginRight, System::SharedPtr<Aspose::Words::Drawing::TextBox> textBox);
    /// <summary>
    /// Checks whether values of properties of an editable range are equal to their expected values.
    /// </summary>
    static void VerifyEditableRange(int32_t expectedId, System::String expectedEditorUser, Aspose::Words::EditorType expectedEditorGroup, System::SharedPtr<Aspose::Words::EditableRange> editableRange);
    /// <summary>
    /// Get File's Encoding.
    /// </summary>
    static System::SharedPtr<System::Text::Encoding> GetEncoding(System::String filename);
    /// <summary>
    /// Converts an image to a byte array.
    /// </summary>
    static System::ArrayPtr<uint8_t> ImageToByteArray(System::String imagePath);
    
private:

    /// <summary>
    /// Checks whether a stream contains a string.
    /// </summary>
    /// <param name="expected">The string we are looking for.</param>
    /// <param name="stream">The stream which, when read from the beginning, should contain the string.</param>
    static void StreamContainsString(System::String expected, System::SharedPtr<System::IO::Stream> stream);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


