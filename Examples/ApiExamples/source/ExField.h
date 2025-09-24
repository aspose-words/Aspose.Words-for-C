#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <system/io/stream.h>
#include <system/collections/list.h>
#include <system/collections/ilist.h>
#include <system/collections/dictionary.h>
#include <system/array.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldUserPromptRespondent.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldUpdatingProgressCallback.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldUpdatingCallback.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldDatabaseProvider.h>
#include <Aspose.Words.Cpp/Model/Fields/IComparisonExpressionEvaluator.h>
#include <Aspose.Words.Cpp/Model/Fields/IBibliographyStylesProvider.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldUpdatingProgressArgs.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Transitional/FieldEQ.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/MergeFieldImageDimensionUnit.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldRef.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldPageRef.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldNoteRef.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldIncludeText.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldTA.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldTime.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/ComparisonExpression.h>
#include <Aspose.Words.Cpp/Model/Fields/ComparisonEvaluationResult.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/GlossaryDocument.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;
using namespace Aspose::Words::Replacing;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExField : public ApiExampleBase
{
    typedef ExField ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    enum class InsertLinkedObjectAs
    {
        Text,
        Unicode,
        Html,
        Rtf,
        Picture,
        Bitmap
    };
    
    
public:

    class OleDbFieldDatabaseProvider : public IFieldDatabaseProvider
    {
        typedef OleDbFieldDatabaseProvider ThisType;
        typedef IFieldDatabaseProvider BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    };
    
    /// <summary>
    /// Document visitor implementation that prints field info.
    /// </summary>
    class FieldVisitor : public DocumentVisitor
    {
        typedef FieldVisitor ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        FieldVisitor();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFieldStart(System::SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart) override;
        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFieldSeparator(System::SharedPtr<Aspose::Words::Fields::FieldSeparator> fieldSeparator) override;
        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitFieldEnd(System::SharedPtr<Aspose::Words::Fields::FieldEnd> fieldEnd) override;
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
    };
    
    class BibliographyStylesProvider : public IBibliographyStylesProvider
    {
        typedef BibliographyStylesProvider ThisType;
        typedef IBibliographyStylesProvider BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    private:
    
        System::SharedPtr<System::IO::Stream> GetStyle(System::String styleFileName) override;
        
    };
    
    /// <summary>
    /// Removes all encountered PRIVATE fields.
    /// </summary>
    class FieldPrivateRemover : public DocumentVisitor
    {
        typedef FieldPrivateRemover ThisType;
        typedef DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        FieldPrivateRemover();
        
        int32_t GetFieldsRemovedCount();
        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// If the node belongs to a PRIVATE field, the entire field is removed.
        /// </summary>
        Aspose::Words::VisitorAction VisitFieldEnd(System::SharedPtr<Aspose::Words::Fields::FieldEnd> fieldEnd) override;
        
    private:
    
        int32_t mFieldsRemovedCount;
        
    };
    
    /// <summary>
    /// Implement this interface if you want to have your own custom methods called during a field update.
    /// </summary>
    class FieldUpdatingCallback : public IFieldUpdatingCallback, public IFieldUpdatingProgressCallback
    {
        typedef FieldUpdatingCallback ThisType;
        typedef IFieldUpdatingCallback BaseType;
        typedef IFieldUpdatingProgressCallback BaseType1;
        
        typedef ::System::BaseTypesInfo<BaseType, BaseType1> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<System::Collections::Generic::IList<System::String>> get_FieldUpdatedCalls() const;
        
        FieldUpdatingCallback();
        
    private:
    
        System::SharedPtr<System::Collections::Generic::IList<System::String>> pr_FieldUpdatedCalls;
        
        /// <summary>
        /// A user defined method that is called just before a field is updated.
        /// </summary>
        void FieldUpdating(System::SharedPtr<Aspose::Words::Fields::Field> field) override;
        /// <summary>
        /// A user defined method that is called just after a field is updated.
        /// </summary>
        void FieldUpdated(System::SharedPtr<Aspose::Words::Fields::Field> field) override;
        void Notify(System::SharedPtr<Aspose::Words::Fields::FieldUpdatingProgressArgs> args) override;
        
    };
    
    
private:

    class InsertTcFieldHandler : public IReplacingCallback
    {
        typedef InsertTcFieldHandler ThisType;
        typedef IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// The display text and switches to use for each TC field. Display name can be an empty String or null.
        /// </summary>
        InsertTcFieldHandler(System::String text, System::String switches);
        
    private:
    
        System::String mFieldText;
        System::String mFieldSwitches;
        
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override;
        
    };
    
    /// <summary>
    /// Prepends text to the default response of an ASK field during a mail merge.
    /// </summary>
    class MyPromptRespondent : public IFieldUserPromptRespondent
    {
        typedef MyPromptRespondent ThisType;
        typedef IFieldUserPromptRespondent BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String Respond(System::String promptText, System::String defaultResponse) override;
        
    };
    
    /// <summary>
    /// Sets the size of all mail merged images to one defined width and height.
    /// </summary>
    class MergedImageResizer : public IFieldMergingCallback
    {
        typedef MergedImageResizer ThisType;
        typedef IFieldMergingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        MergedImageResizer(double imageWidth, double imageHeight, Aspose::Words::Fields::MergeFieldImageDimensionUnit unit);
        
        void FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> e) override;
        void ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args) override;
        
    private:
    
        double mImageWidth;
        double mImageHeight;
        Aspose::Words::Fields::MergeFieldImageDimensionUnit mUnit;
        
    };
    
    /// <summary>
    /// Contains a dictionary that maps names of images to local system filenames that contain these images.
    /// If a mail merge data source uses one of the dictionary's names to refer to an image,
    /// this callback will pass the respective filename to the merge destination.
    /// </summary>
    class ImageFilenameCallback : public IFieldMergingCallback
    {
        typedef ImageFilenameCallback ThisType;
        typedef IFieldMergingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        ImageFilenameCallback();
        
    private:
    
        System::SharedPtr<System::Collections::Generic::Dictionary<System::String, System::String>> mImageFilenames;
        
        void FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args) override;
        void ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args) override;
        
    };
    
    /// <summary>
    /// Prepends a line to the default response of every FILLIN field during a mail merge.
    /// </summary>
    class PromptRespondent : public IFieldUserPromptRespondent
    {
        typedef PromptRespondent ThisType;
        typedef IFieldUserPromptRespondent BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String Respond(System::String promptText, System::String defaultResponse) override;
        
    };
    
    /// <summary>
    /// Comparison expressions evaluation for the FieldIf and FieldCompare.
    /// </summary>
    class ComparisonExpressionEvaluator : public IComparisonExpressionEvaluator
    {
        typedef ComparisonExpressionEvaluator ThisType;
        typedef IComparisonExpressionEvaluator BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        ComparisonExpressionEvaluator(System::SharedPtr<Aspose::Words::Fields::ComparisonEvaluationResult> result);
        
        System::SharedPtr<Aspose::Words::Fields::ComparisonEvaluationResult> Evaluate(System::SharedPtr<Aspose::Words::Fields::Field> field, System::SharedPtr<Aspose::Words::Fields::ComparisonExpression> expression) override;
        System::SharedPtr<Aspose::Words::ApiExamples::ExField::ComparisonExpressionEvaluator> AssertInvocationsCount(int32_t expected);
        System::SharedPtr<Aspose::Words::ApiExamples::ExField::ComparisonExpressionEvaluator> AssertInvocationArguments(int32_t invocationIndex, System::String expectedLeftExpression, System::String expectedComparisonOperator, System::String expectedRightExpression);
        
    protected:
    
        virtual ~ComparisonExpressionEvaluator();
        
    private:
    
        System::SharedPtr<Aspose::Words::Fields::ComparisonEvaluationResult> mResult;
        System::SharedPtr<System::Collections::Generic::List<System::ArrayPtr<System::String>>> mInvocations;
        
    };
    
    
public:

    void GetFieldFromDocument();
    void GetFieldData();
    void GetFieldCode();
    void DisplayResult();
    void CreateWithFieldBuilder();
    void RevNum();
    void InsertFieldNone();
    void InsertTcField();
    void InsertTcFieldsAtText();
    void FieldLocale();
    void UpdateDirtyFields(bool updateDirtyFields);
    void InsertFieldWithFieldBuilderException();
    void PreserveIncludePicture(bool preserveIncludePictureField);
    void FieldFormat();
    void Unlink();
    void UnlinkAllFieldsInRange();
    void UnlinkSingleField();
    void UpdateTocPageNumbers();
    void FieldAdvance();
    void FieldAddressBlock();
    void FieldCollection();
    void RemoveFields();
    void FieldCompare();
    void FieldIf();
    void FieldAutoNum();
    void FieldAutoNumLgl();
    void FieldAutoNumOut();
    void FieldAutoText();
    void FieldAutoTextList();
    void FieldListNum();
    void FieldToc();
    /// <summary>
    /// Start a new page and insert a paragraph of a specified style.
    /// </summary>
    void InsertNewPageWithHeading(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String captionText, System::String styleName);
    void FieldTocEntryIdentifier();
    /// <summary>
    /// Use a document builder to insert a TC field.
    /// </summary>
    void InsertTocEntry(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String text, System::String typeIdentifier, System::String entryLevel);
    void TocSeqPrefix();
    void TocSeqNumbering();
    void TocSeqBookmark();
    void ChangeBibliographyStyles();
    void FieldData();
    void FieldInclude();
    void FieldIncludePicture();
    /// <summary>
    /// Use a document builder to insert an INCLUDETEXT field with custom properties.
    /// </summary>
    System::SharedPtr<Aspose::Words::Fields::FieldIncludeText> CreateFieldIncludeText(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String sourceFullName, bool lockFields, System::String mimeType, System::String textConverter, System::String encoding);
    void FieldHyperlink();
    void FieldIndexFilter();
    void FieldIndexFormatting();
    void FieldIndexSequence();
    void FieldIndexPageNumberSeparator();
    void FieldIndexPageRangeBookmark();
    void FieldIndexCrossReferenceSeparator();
    void FieldIndexSubheading(bool runSubentriesOnTheSameLine);
    void FieldIndexYomi(bool sortEntriesUsingYomi);
    void FieldBarcode();
    void FieldDisplayBarcode();
    void FieldLinkedObjectsAsText(Aspose::Words::ApiExamples::ExField::InsertLinkedObjectAs insertLinkedObjectAs);
    void FieldLinkedObjectsAsImage(Aspose::Words::ApiExamples::ExField::InsertLinkedObjectAs insertLinkedObjectAs);
    void FieldUserAddress();
    void FieldUserInitials();
    void FieldUserName();
    void FieldStyleRefParagraphNumbers();
    void FieldDate();
    void FieldCreateDate();
    void FieldSaveDate();
    void FieldBuilder();
    void FieldAuthor();
    void FieldDocVariable();
    void FieldSubject();
    void FieldComments();
    void FieldFileSize();
    void FieldGoToButton();
    void FieldFillIn();
    void FieldInfo();
    void FieldMacroButton();
    void FieldKeywords();
    void FieldNum();
    void FieldPrint();
    void FieldPrintDate();
    void FieldQuote();
    /// <summary>
    /// Uses a document builder to insert MERGEFIELDs for a data source that contains columns named "Courtesy Title", "First Name" and "Last Name".
    /// </summary>
    void InsertMergeFields(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String firstFieldTextBefore);
    /// <summary>
    /// Uses a document builder to insert a MERRGEFIELD with specified properties.
    /// </summary>
    void InsertMergeField(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String fieldName, System::String textBefore, System::String textAfter);
    void FieldNoteRef();
    void NoteRef();
    void FieldPageRef();
    void FieldRef();
    void FieldRD();
    void FieldSetRef();
    void FieldTemplate();
    void FieldSymbol();
    void FieldTitle();
    void FieldTOA();
    void FieldAddIn();
    void FieldEditTime();
    void FieldEQ();
    void FieldEQAsOfficeMath();
    void FieldForms();
    void FieldFormula();
    void FieldLastSavedBy();
    void FieldOcx();
    void FieldPrivate();
    void FieldSection();
    void FieldTime();
    void BidiOutline();
    void Legacy();
    void SetFieldIndexFormat();
    void ConditionEvaluationExtensionPoint(System::String fieldCode, int8_t comparisonResult, System::String comparisonError, System::String expectedResult);
    void ComparisonExpressionEvaluatorNestedFields();
    void ComparisonExpressionEvaluatorHeaderFooterFields();
    void FieldUpdatingCallbackTest();
    void BibliographySources();
    void BibliographyPersons();
    void CaptionlessTableOfFiguresLabel();
    
protected:

    static void RemoveSequence(System::SharedPtr<Aspose::Words::Node> start, System::SharedPtr<Aspose::Words::Node> end);
    void TestFieldCollection(System::String fieldVisitorText);
    /// <summary>
    /// Uses a document builder to insert a clause numbered by an AUTONUMLGL field.
    /// </summary>
    static void InsertNumberedClause(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String heading, System::String contents, Aspose::Words::StyleIdentifier headingStyle);
    void TestFieldAutoNumLgl(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Create an AutoText-type building block and add it to a glossary document.
    /// </summary>
    static void AppendAutoTextEntry(System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> glossaryDoc, System::String name, System::String contents);
    void TestFieldAutoTextList(System::SharedPtr<Aspose::Words::Document> doc);
    void TestFieldToc(System::SharedPtr<Aspose::Words::Document> doc);
    void TestFieldTocEntryIdentifier(System::SharedPtr<Aspose::Words::Document> doc);
    void TestFieldIncludeText(System::SharedPtr<Aspose::Words::Document> doc);
    void TestMergeFieldImageDimension(System::SharedPtr<Aspose::Words::Document> doc);
    void TestMergeFieldImages(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Use a document builder to insert a LINK field and set its properties according to parameters.
    /// </summary>
    static void InsertFieldLink(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, Aspose::Words::ApiExamples::ExField::InsertLinkedObjectAs insertLinkedObjectAs, System::String progId, System::String sourceFullName, System::String sourceItem, bool shouldAutoUpdate);
    /// <summary>
    /// Use a document builder to insert a DDE field, and set its properties according to parameters.
    /// </summary>
    static void InsertFieldDde(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, Aspose::Words::ApiExamples::ExField::InsertLinkedObjectAs insertLinkedObjectAs, System::String progId, System::String sourceFullName, System::String sourceItem, bool isLinked, bool shouldAutoUpdate);
    /// <summary>
    /// Use a document builder to insert a DDEAUTO, field and set its properties according to parameters.
    /// </summary>
    static void InsertFieldDdeAuto(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, Aspose::Words::ApiExamples::ExField::InsertLinkedObjectAs insertLinkedObjectAs, System::String progId, System::String sourceFullName, System::String sourceItem, bool isLinked);
    void TestFieldFillIn(System::SharedPtr<Aspose::Words::Document> doc);
    void TestFieldNext(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Uses a document builder to insert a NOTEREF field with specified properties.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldNoteRef> InsertFieldNoteRef(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, bool insertHyperlink, bool insertRelativePosition, bool insertReferenceMark, System::String textBefore);
    /// <summary>
    /// Uses a document builder to insert a named bookmark with a footnote at the end.
    /// </summary>
    static void InsertBookmarkWithFootnote(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, System::String bookmarkText, System::String footnoteText);
    void TestNoteRef(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Uses a document builder to insert a PAGEREF field and sets its properties.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldPageRef> InsertFieldPageRef(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, bool insertHyperlink, bool insertRelativePosition, System::String textBefore);
    /// <summary>
    /// Uses a document builder to insert a named bookmark.
    /// </summary>
    static void InsertAndNameBookmark(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName);
    void TestPageRef(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Get the document builder to insert a REF field, reference a bookmark with it, and add text before and after it.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldRef> InsertFieldRef(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, System::String textBefore, System::String textAfter);
    void TestFieldRef(System::SharedPtr<Aspose::Words::Document> doc);
    static System::SharedPtr<Aspose::Words::Fields::FieldTA> InsertToaEntry(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String entryCategory, System::String longCitation);
    void TestFieldTOA(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Use a document builder to insert an EQ field, set its arguments and start a new paragraph.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldEQ> InsertFieldEQ(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String args);
    void TestFieldEQ(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Use a document builder to insert a TIME field, insert a new paragraph and return the field.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldTime> InsertFieldTime(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String format);
    void TestFieldTime(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


