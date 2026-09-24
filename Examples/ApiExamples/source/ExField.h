#pragma once

#include <system/text/string_builder.h>
#include <system/io/stream.h>
#include <system/collections/list.h>
#include <system/collections/ilist.h>
#include <system/collections/dictionary.h>
#include <system/array.h>
#include <gtest/gtest.h>
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
        // LinkedObjectAsText
        Text,
        Unicode,
        Html,
        Rtf,
        // LinkedObjectAsImage
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
    //ExStart
    //ExFor:FieldCollection
    //ExFor:FieldCollection.Count
    //ExFor:FieldCollection.GetEnumerator
    //ExFor:FieldStart
    //ExFor:FieldStart.Accept(DocumentVisitor)
    //ExFor:FieldSeparator
    //ExFor:FieldSeparator.Accept(DocumentVisitor)
    //ExFor:FieldEnd
    //ExFor:FieldEnd.Accept(DocumentVisitor)
    //ExFor:FieldEnd.HasSeparator
    //ExFor:Field.End
    //ExFor:Field.Separator
    //ExFor:Field.Start
    //ExSummary:Shows how to work with a collection of fields.
    void FieldCollection();
    void RemoveFields();
    void FieldCompare();
    void FieldIf();
    void FieldAutoNum();
    //ExStart
    //ExFor:FieldAutoNumLgl
    //ExFor:FieldAutoNumLgl.RemoveTrailingPeriod
    //ExFor:FieldAutoNumLgl.SeparatorCharacter
    //ExSummary:Shows how to organize a document using AUTONUMLGL fields.
    void FieldAutoNumLgl();
    void FieldAutoNumOut();
    void FieldAutoText();
    //ExStart
    //ExFor:FieldAutoTextList
    //ExFor:FieldAutoTextList.EntryName
    //ExFor:FieldAutoTextList.ListStyle
    //ExFor:FieldAutoTextList.ScreenTip
    //ExSummary:Shows how to use an AUTOTEXTLIST field to select from a list of AutoText entries.
    void FieldAutoTextList();
    void FieldListNum();
    //ExStart
    //ExFor:FieldToc
    //ExFor:FieldToc.BookmarkName
    //ExFor:FieldToc.CustomStyles
    //ExFor:FieldToc.EntrySeparator
    //ExFor:FieldToc.HeadingLevelRange
    //ExFor:FieldToc.HideInWebLayout
    //ExFor:FieldToc.InsertHyperlinks
    //ExFor:FieldToc.PageNumberOmittingLevelRange
    //ExFor:FieldToc.PreserveLineBreaks
    //ExFor:FieldToc.PreserveTabs
    //ExFor:FieldToc.UpdatePageNumbers
    //ExFor:FieldToc.UseParagraphOutlineLevel
    //ExFor:FieldOptions.CustomTocStyleSeparator
    //ExSummary:Shows how to insert a TOC, and populate it with entries based on heading styles.
    void FieldToc();
    /// <summary>
    /// Start a new page and insert a paragraph of a specified style.
    /// </summary>
    void InsertNewPageWithHeading(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String captionText, System::String styleName);
    //ExStart
    //ExFor:FieldToc.EntryIdentifier
    //ExFor:FieldToc.EntryLevelRange
    //ExFor:FieldTC
    //ExFor:FieldTC.OmitPageNumber
    //ExFor:FieldTC.Text
    //ExFor:FieldTC.TypeIdentifier
    //ExFor:FieldTC.EntryLevel
    //ExSummary:Shows how to insert a TOC field, and filter which TC fields end up as entries.
    void FieldTocEntryIdentifier();
    /// <summary>
    /// Use a document builder to insert a TC field.
    /// </summary>
    void InsertTocEntry(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String text, System::String typeIdentifier, System::String entryLevel);
    void TocSeqPrefix();
    void TocSeqNumbering();
    void TocSeqBookmark();
    //ExStart
    //ExFor:Bibliography.BibliographyStyle
    //ExFor:IBibliographyStylesProvider
    //ExFor:IBibliographyStylesProvider.GetStyle(String)
    //ExFor:FieldOptions.BibliographyStylesProvider
    //ExSummary:Shows how to override built-in styles or provide custom one.
    void ChangeBibliographyStyles();
    //ExEnd
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
    //ExStart
    //ExFor:FieldLink
    //ExFor:FieldLink.AutoUpdate
    //ExFor:FieldLink.FormatUpdateType
    //ExFor:FieldLink.InsertAsBitmap
    //ExFor:FieldLink.InsertAsHtml
    //ExFor:FieldLink.InsertAsPicture
    //ExFor:FieldLink.InsertAsRtf
    //ExFor:FieldLink.InsertAsText
    //ExFor:FieldLink.InsertAsUnicode
    //ExFor:FieldLink.IsLinked
    //ExFor:FieldLink.ProgId
    //ExFor:FieldLink.SourceFullName
    //ExFor:FieldLink.SourceItem
    //ExFor:FieldDde
    //ExFor:FieldDde.AutoUpdate
    //ExFor:FieldDde.InsertAsBitmap
    //ExFor:FieldDde.InsertAsHtml
    //ExFor:FieldDde.InsertAsPicture
    //ExFor:FieldDde.InsertAsRtf
    //ExFor:FieldDde.InsertAsText
    //ExFor:FieldDde.InsertAsUnicode
    //ExFor:FieldDde.IsLinked
    //ExFor:FieldDde.ProgId
    //ExFor:FieldDde.SourceFullName
    //ExFor:FieldDde.SourceItem
    //ExFor:FieldDdeAuto
    //ExFor:FieldDdeAuto.InsertAsBitmap
    //ExFor:FieldDdeAuto.InsertAsHtml
    //ExFor:FieldDdeAuto.InsertAsPicture
    //ExFor:FieldDdeAuto.InsertAsRtf
    //ExFor:FieldDdeAuto.InsertAsText
    //ExFor:FieldDdeAuto.InsertAsUnicode
    //ExFor:FieldDdeAuto.IsLinked
    //ExFor:FieldDdeAuto.ProgId
    //ExFor:FieldDdeAuto.SourceFullName
    //ExFor:FieldDdeAuto.SourceItem
    //ExSummary:Shows how to use various field types to link to other documents in the local file system, and display their contents.
    void FieldLinkedObjectsAsText(Aspose::Words::ApiExamples::ExField::InsertLinkedObjectAs insertLinkedObjectAs);
    void FieldLinkedObjectsAsImage(Aspose::Words::ApiExamples::ExField::InsertLinkedObjectAs insertLinkedObjectAs);
    //ExEnd
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
    //ExStart
    //ExFor:FieldNoteRef
    //ExFor:FieldNoteRef.BookmarkName
    //ExFor:FieldNoteRef.InsertHyperlink
    //ExFor:FieldNoteRef.InsertReferenceMark
    //ExFor:FieldNoteRef.InsertRelativePosition
    //ExSummary:Shows to insert NOTEREF fields, and modify their appearance.
    void FieldNoteRef();
    void NoteRef();
    //ExStart
    //ExFor:FieldPageRef
    //ExFor:FieldPageRef.BookmarkName
    //ExFor:FieldPageRef.InsertHyperlink
    //ExFor:FieldPageRef.InsertRelativePosition
    //ExSummary:Shows to insert PAGEREF fields to display the relative location of bookmarks.
    void FieldPageRef();
    //ExStart
    //ExFor:FieldRef
    //ExFor:FieldRef.BookmarkName
    //ExFor:FieldRef.IncludeNoteOrComment
    //ExFor:FieldRef.InsertHyperlink
    //ExFor:FieldRef.InsertParagraphNumber
    //ExFor:FieldRef.InsertParagraphNumberInFullContext
    //ExFor:FieldRef.InsertParagraphNumberInRelativeContext
    //ExFor:FieldRef.InsertRelativePosition
    //ExFor:FieldRef.NumberSeparator
    //ExFor:FieldRef.SuppressNonDelimiters
    //ExSummary:Shows how to insert REF fields to reference bookmarks.
    void FieldRef();
    void FieldRD();
    void FieldSetRef();
    void FieldTemplate();
    void FieldSymbol();
    void FieldTitle();
    //ExStart
    //ExFor:FieldToa
    //ExFor:FieldToa.BookmarkName
    //ExFor:FieldToa.EntryCategory
    //ExFor:FieldToa.EntrySeparator
    //ExFor:FieldToa.PageNumberListSeparator
    //ExFor:FieldToa.PageRangeSeparator
    //ExFor:FieldToa.RemoveEntryFormatting
    //ExFor:FieldToa.SequenceName
    //ExFor:FieldToa.SequenceSeparator
    //ExFor:FieldToa.UseHeading
    //ExFor:FieldToa.UsePassim
    //ExFor:FieldTA
    //ExFor:FieldTA.EntryCategory
    //ExFor:FieldTA.IsBold
    //ExFor:FieldTA.IsItalic
    //ExFor:FieldTA.LongCitation
    //ExFor:FieldTA.PageRangeBookmarkName
    //ExFor:FieldTA.ShortCitation
    //ExSummary:Shows how to build and customize a table of authorities using TOA and TA fields.
    void FieldTOA();
    void FieldAddIn();
    void FieldEditTime();
    //ExStart
    //ExFor:FieldEQ
    //ExSummary:Shows how to use the EQ field to display a variety of mathematical equations.
    void FieldEQ();
    void FieldEQAsOfficeMath();
    void FieldForms();
    void FieldFormula();
    void FieldLastSavedBy();
    void FieldOcx();
    //ExStart
    //ExFor:Field.Remove
    //ExFor:FieldPrivate
    //ExSummary:Shows how to process PRIVATE fields.
    void FieldPrivate();
    //ExEnd
    void FieldSection();
    //ExStart
    //ExFor:FieldTime
    //ExSummary:Shows how to display the current time using the TIME field.
    void FieldTime();
    void BidiOutline();
    void Legacy();
    void SetFieldIndexFormat();
    //ExStart
    //ExFor:ComparisonEvaluationResult.#ctor(bool)
    //ExFor:ComparisonEvaluationResult.#ctor(string)
    //ExFor:ComparisonEvaluationResult
    //ExFor:ComparisonEvaluationResult.ErrorMessage
    //ExFor:ComparisonEvaluationResult.Result
    //ExFor:ComparisonExpression
    //ExFor:ComparisonExpression.LeftExpression
    //ExFor:ComparisonExpression.ComparisonOperator
    //ExFor:ComparisonExpression.RightExpression
    //ExFor:FieldOptions.ComparisonExpressionEvaluator
    //ExFor:IComparisonExpressionEvaluator
    //ExFor:IComparisonExpressionEvaluator.Evaluate(Field,ComparisonExpression)
    //ExSummary:Shows how to implement custom evaluation for the IF and COMPARE fields.
    void ConditionEvaluationExtensionPoint(System::String fieldCode, int8_t comparisonResult, System::String comparisonError, System::String expectedResult);
    //ExEnd
    void ComparisonExpressionEvaluatorNestedFields();
    void ComparisonExpressionEvaluatorHeaderFooterFields();
    //ExStart
    //ExFor:FieldOptions.FieldUpdatingCallback
    //ExFor:FieldOptions.FieldUpdatingProgressCallback
    //ExFor:IFieldUpdatingCallback
    //ExFor:IFieldUpdatingProgressCallback
    //ExFor:IFieldUpdatingProgressCallback.Notify(FieldUpdatingProgressArgs)
    //ExFor:FieldUpdatingProgressArgs
    //ExFor:FieldUpdatingProgressArgs.UpdateCompleted
    //ExFor:FieldUpdatingProgressArgs.TotalFieldsCount
    //ExFor:FieldUpdatingProgressArgs.UpdatedFieldsCount
    //ExFor:IFieldUpdatingCallback.FieldUpdating(Field)
    //ExFor:IFieldUpdatingCallback.FieldUpdated(Field)
    //ExSummary:Shows how to use callback methods during a field update.
    void FieldUpdatingCallbackTest();
    //ExEnd
    void BibliographySources();
    void BibliographyPersons();
    void CaptionlessTableOfFiguresLabel();
    
protected:

    static void RemoveSequence(System::SharedPtr<Aspose::Words::Node> start, System::SharedPtr<Aspose::Words::Node> end);
    //ExEnd
    void TestFieldCollection(System::String fieldVisitorText);
    /// <summary>
    /// Uses a document builder to insert a clause numbered by an AUTONUMLGL field.
    /// </summary>
    static void InsertNumberedClause(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String heading, System::String contents, Aspose::Words::StyleIdentifier headingStyle);
    //ExEnd
    void TestFieldAutoNumLgl(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Create an AutoText-type building block and add it to a glossary document.
    /// </summary>
    static void AppendAutoTextEntry(System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> glossaryDoc, System::String name, System::String contents);
    //ExEnd
    void TestFieldAutoTextList(System::SharedPtr<Aspose::Words::Document> doc);
    //ExEnd
    void TestFieldToc(System::SharedPtr<Aspose::Words::Document> doc);
    //ExEnd
    void TestFieldTocEntryIdentifier(System::SharedPtr<Aspose::Words::Document> doc);
    //ExEnd
    void TestFieldIncludeText(System::SharedPtr<Aspose::Words::Document> doc);
    //ExEnd
    void TestMergeFieldImageDimension(System::SharedPtr<Aspose::Words::Document> doc);
    //ExEnd
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
    //ExEnd
    void TestFieldFillIn(System::SharedPtr<Aspose::Words::Document> doc);
    //ExEnd
    void TestFieldNext(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Uses a document builder to insert a NOTEREF field with specified properties.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldNoteRef> InsertFieldNoteRef(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, bool insertHyperlink, bool insertRelativePosition, bool insertReferenceMark, System::String textBefore);
    /// <summary>
    /// Uses a document builder to insert a named bookmark with a footnote at the end.
    /// </summary>
    static void InsertBookmarkWithFootnote(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, System::String bookmarkText, System::String footnoteText);
    //ExEnd
    void TestNoteRef(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Uses a document builder to insert a PAGEREF field and sets its properties.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldPageRef> InsertFieldPageRef(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, bool insertHyperlink, bool insertRelativePosition, System::String textBefore);
    /// <summary>
    /// Uses a document builder to insert a named bookmark.
    /// </summary>
    static void InsertAndNameBookmark(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName);
    //ExEnd
    void TestPageRef(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Get the document builder to insert a REF field, reference a bookmark with it, and add text before and after it.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldRef> InsertFieldRef(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, System::String textBefore, System::String textAfter);
    //ExEnd
    void TestFieldRef(System::SharedPtr<Aspose::Words::Document> doc);
    static System::SharedPtr<Aspose::Words::Fields::FieldTA> InsertToaEntry(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String entryCategory, System::String longCitation);
    //ExEnd
    void TestFieldTOA(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Use a document builder to insert an EQ field, set its arguments and start a new paragraph.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldEQ> InsertFieldEQ(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String args);
    //ExEnd
    void TestFieldEQ(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Use a document builder to insert a TIME field, insert a new paragraph and return the field.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldTime> InsertFieldTime(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String format);
    //ExEnd
    void TestFieldTime(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


