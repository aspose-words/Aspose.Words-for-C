#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: System.Data is not supported
// CPPDEFECT: Aspose.BarCode is not supported

#include <system/text/string_builder.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/collections/dictionary.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldUserPromptRespondent.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Replacing { enum class ReplaceAction; } } }
namespace Aspose { namespace Words { namespace Replacing { class ReplacingArgs; } } }
namespace Aspose { namespace Words { class Node; } }
namespace Aspose { namespace Words { enum class VisitorAction; } }
namespace Aspose { namespace Words { namespace Fields { class FieldStart; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldSeparator; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldEnd; } } }
namespace Aspose { namespace Words { class DocumentBuilder; } }
namespace Aspose { namespace Words { enum class StyleIdentifier; } }
namespace Aspose { namespace Words { class Document; } }
namespace Aspose { namespace Words { namespace BuildingBlocks { class GlossaryDocument; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldIncludeText; } } }
namespace Aspose { namespace Words { namespace MailMerging { class FieldMergingArgs; } } }
namespace Aspose { namespace Words { namespace MailMerging { class ImageFieldMergingArgs; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldNoteRef; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldPageRef; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldRef; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldTA; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldEQ; } } }
namespace Aspose { namespace Words { namespace Fields { class FieldTime; } } }

namespace ApiExamples {

class ExField : public ApiExampleBase
{
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

    /// <summary>
    /// Document visitor implementation that prints field info.
    /// </summary>
    class FieldVisitor : public Aspose::Words::DocumentVisitor
    {
        typedef FieldVisitor ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
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
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
    };
    
    /// <summary>
    /// Visitor implementation that removes all PRIVATE fields that it encounters.
    /// </summary>
    class FieldPrivateRemover : public Aspose::Words::DocumentVisitor
    {
        typedef FieldPrivateRemover ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
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
    
    
private:

    class InsertTcFieldHandler : public Aspose::Words::Replacing::IReplacingCallback
    {
        typedef InsertTcFieldHandler ThisType;
        typedef Aspose::Words::Replacing::IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// The display text and switches to use for each TC field. Display name can be an empty String or null.
        /// </summary>
        InsertTcFieldHandler(System::String text, System::String switches);
        
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override;
        
    private:
    
        System::String mFieldText;
        System::String mFieldSwitches;
        
    };
    
    /// <summary>
    /// Image merging callback that pairs image shorthand names with filenames.
    /// </summary>
    class ImageFilenameCallback : public Aspose::Words::MailMerging::IFieldMergingCallback
    {
        typedef ImageFilenameCallback ThisType;
        typedef Aspose::Words::MailMerging::IFieldMergingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        ImageFilenameCallback();
        
        void FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args) override;
        void ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Collections::Generic::Dictionary<System::String, System::String>> mImageFilenames;
        
    };
    
    /// <summary>
    /// IFieldUserPromptRespondent implementation that appends a line to the default response of an FILLIN field during a mail merge.
    /// </summary>
    class PromptRespondent : public Aspose::Words::Fields::IFieldUserPromptRespondent
    {
        typedef PromptRespondent ThisType;
        typedef Aspose::Words::Fields::IFieldUserPromptRespondent BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String Respond(System::String promptText, System::String defaultResponse) override;
        
    };
    
    
public:

    void GetFieldFromDocument();
    void GetFieldCode();
    void FieldDisplayResult();
    void CreateWithFieldBuilder();
    void CreateRevNumFieldByDocumentBuilder();
    void CreateInfoFieldWithFieldBuilder();
    void CreateInfoFieldWithDocumentBuilder();
    void GetFieldFromFieldCollection();
    void InsertFieldNone();
    void InsertTcField();
    void FieldLocale();
    void ChangeLocale();
    void RemoveTocFromDocument();
    void InsertTcFieldsAtText();
    void UpdateDirtyFields(bool doUpdateDirtyFields);
    void InsertFieldWithFieldBuilderException();
    void UpdateFieldIgnoringMergeFormat();
    void FieldFormat();
    void UnlinkAllFieldsInDocument();
    void UnlinkAllFieldsInRange();
    void UnlinkSingleField();
    void UpdateTocPageNumbers();
    void DropDownItemCollection();
    void FieldAdvance();
    void FieldAddressBlock();
    void FieldCollection();
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
    /// Insert a table of contents entry via a document builder.
    /// </summary>
    void InsertTocEntry(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String text, System::String typeIdentifier, System::String entryLevel);
    void TocSeqPrefix();
    void TocSeqNumbering();
    void TocSeqBookmark();
    void FieldCitation();
    void FieldData();
    void FieldInclude();
    void FieldIncludePicture();
    void FieldIncludeText();
    /// <summary>
    /// Use a document builder to insert an INCLUDETEXT field and set its properties.
    /// </summary>
    System::SharedPtr<Aspose::Words::Fields::FieldIncludeText> CreateFieldIncludeText(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String sourceFullName, bool lockFields, System::String mimeType, System::String textConverter, System::String encoding);
    void FieldHyperlink();
    void FieldIndexFilter();
    void FieldIndexFormatting();
    void FieldIndexSequence();
    void FieldIndexPageNumberSeparator();
    void FieldIndexPageRangeBookmark();
    void FieldIndexCrossReferenceSeparator();
    void FieldIndexSubheading(bool doRunSubentriesOnTheSameLine);
    void FieldIndexYomi(bool doSortEntriesUsingYomi);
    void FieldBarcode();
    void FieldDisplayBarcode();
    void FieldLinkedObjectsAsText(ExField::InsertLinkedObjectAs insertLinkedObjectAs);
    void FieldLinkedObjectsAsImage(ExField::InsertLinkedObjectAs insertLinkedObjectAs);
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
    void FieldNoteRef();
    void FootnoteRef();
    void FieldPageRef();
    void FieldRef();
    void FieldRD();
    void FieldSet();
    void FieldTemplate();
    void FieldSymbol();
    void FieldTitle();
    void FieldTOA();
    void FieldAddIn();
    void FieldEditTime();
    void FieldEQ();
    void FieldForms();
    void FieldFormula();
    void FieldLastSavedBy();
    void FieldOcx();
    void FieldPrivate();
    void FieldSection();
    void FieldTime();
    void BidiOutline();
    void Legacy();
    
protected:

    static void RemoveSequence(System::SharedPtr<Aspose::Words::Node> start, System::SharedPtr<Aspose::Words::Node> end);
    /// <summary>
    /// Get a document builder to insert a clause numbered by an AUTONUMLGL field.
    /// </summary>
    static void InsertNumberedClause(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String heading, System::String contents, Aspose::Words::StyleIdentifier headingStyle);
    void TestFieldAutoNumLgl(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Create an AutoText entry and add it to a glossary document.
    /// </summary>
    static void AppendAutoTextEntry(System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> glossaryDoc, System::String name, System::String contents);
    void TestFieldToc(System::SharedPtr<Aspose::Words::Document> doc);
    void TestFieldTocEntryIdentifier(System::SharedPtr<Aspose::Words::Document> doc);
    void TestFieldIncludeText(System::SharedPtr<Aspose::Words::Document> doc);
    void TestMergeFieldImages(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Use a document builder to insert a LINK field and set its properties according to parameters.
    /// </summary>
    static void InsertFieldLink(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, ExField::InsertLinkedObjectAs insertLinkedObjectAs, System::String progId, System::String sourceFullName, System::String sourceItem, bool shouldAutoUpdate);
    /// <summary>
    /// Use a document builder to insert a DDE field and set its properties according to parameters.
    /// </summary>
    static void InsertFieldDde(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, ExField::InsertLinkedObjectAs insertLinkedObjectAs, System::String progId, System::String sourceFullName, System::String sourceItem, bool isLinked, bool shouldAutoUpdate);
    /// <summary>
    /// Use a document builder to insert a DDEAUTO field and set its properties according to parameters.
    /// </summary>
    static void InsertFieldDdeAuto(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, ExField::InsertLinkedObjectAs insertLinkedObjectAs, System::String progId, System::String sourceFullName, System::String sourceItem, bool isLinked);
    void TestFieldFillIn(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Uses a document builder to insert a NOTEREF field and sets its attributes.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldNoteRef> InsertFieldNoteRef(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, bool insertHyperlink, bool insertRelativePosition, bool insertReferenceMark, System::String textBefore);
    /// <summary>
    /// Uses a document builder to insert a named bookmark with a footnote at the end.
    /// </summary>
    static void InsertBookmarkWithFootnote(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, System::String bookmarkText, System::String footnoteText);
    void TestNoteRef(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Uses a document builder to insert a PAGEREF field and sets its attributes.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldPageRef> InsertFieldPageRef(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, bool insertHyperlink, bool insertRelativePosition, System::String textBefore);
    /// <summary>
    /// Uses a document builder to insert a named bookmark.
    /// </summary>
    static void InsertAndNameBookmark(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName);
    void TestPageRef(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Get the document builder to insert a REF field, reference a bookmark with it, and add text before and after.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldRef> InsertFieldRef(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, System::String textBefore, System::String textAfter);
    void TestFieldRef(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Get a builder to insert a TA field, specifying its long citation and category,
    /// then insert a page break and return the field we created.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldTA> InsertToaEntry(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String entryCategory, System::String longCitation);
    void TestFieldTOA(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Use a document builder to insert an EQ field, set its arguments and start a new paragraph.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldEQ> InsertFieldEQ(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String args);
    void TestFieldEQ(System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Use a document builder to insert a TIME field, insert a new paragraph and return the field
    /// </summary>
    static System::SharedPtr<Aspose::Words::Fields::FieldTime> InsertFieldTime(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String format);
    void TestFieldTime(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples


