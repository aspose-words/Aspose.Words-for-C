#include "stdafx.h"
#include <iostream>

#include "examples.h"

int main()
{
    std::cout << "Examples:" << std::endl << "=====================================================" << std::endl << std::endl;

    // Uncomment the one you want to try out

    // =====================================================
    // =====================================================
    // Quick Start
    // =====================================================
    // =====================================================
    AppendDocuments();
    ApplyLicense();
    FindAndReplace();
    HelloWorld();
    WorkingWithNodes();
    UpdateFields();
    ApplyLicenseFromFile();
    ApplyLicenseFromStream();

    // =====================================================
    // =====================================================
    // Loading and Saving
    // =====================================================
    // =====================================================
    LoadAndSaveToDisk();
    LoadAndSaveToStream();
    CreateDocument();
    ConvertDocumentToByte();
    LoadTxt();
    DetectDocumentSignatures();
    WorkingWithTxt();
    WorkingWithRTF();
    Load_Options();

    // =====================================================
    // =====================================================
    // Programming with Documents
    // =====================================================
    // =====================================================
    
    // Joining and Appending
    // =====================================================
    SimpleAppendDocument();
    KeepSourceFormatting();
    UseDestinationStyles();
    JoinContinuous();
    JoinNewPage();
    RestartPageNumbering();
    LinkHeadersFooters();
    UnlinkHeadersFooters();
    RemoveSourceHeadersFooters();
    DifferentPageSetup();
    ListUseDestinationStyles();
    ListKeepSourceFormatting();
    KeepSourceTogether();
    BaseDocument();
    AppendDocumentManually();
    PrependDocument();

    // Find and Replace
    // =====================================================
    FindAndHighlight();
    ReplaceWithString();
    ReplaceWithRegex();
    ReplaceWithEvaluator();
    FindReplaceUsingMetaCharacters();
    ReplaceTextWithField();

    // Bookmarks
    // =====================================================
    CopyBookmarkedText();
    UntangleRowBookmarks();
    BookmarkTable();
    BookmarkNameAndText();
    AccessBookmarks();
    CreateBookmark();

    // Shapes
    // =====================================================
    WorkingWithShapes();

    // Comments
    // =====================================================
    ProcessComments();
    AddComments();
    AnchorComment();
    CommentReply();

    // ConvertUtil
    // =====================================================
    UtilityClasses();

    // Document
    // =====================================================
    ExtractContentBetweenParagraphs();
    ExtractContentBetweenBlockLevelNodes();
    ExtractContentBetweenParagraphStyles();
    ExtractContentBetweenRuns();
    ExtractContentBetweenBookmark();
    ExtractContentBetweenCommentRange();
    RemoveBreaks();
    CloningDocument();
    ProtectDocument();
    AccessStyles();
    GetVariables();
    SetViewOption();
    CreateHeaderFooterUsingDocBuilder();
    RemoveFooters();
    AddGroupShapeToDocument();
    DocProperties();
    WriteAndFont();
    DocumentBuilderInsertParagraph();
    DocumentBuilderBuildTable();
    DocumentBuilderInsertBreak();
    DocumentBuilderInsertImage();
    DocumentBuilderInsertBookmark();
    DocumentBuilderInsertElements();
    DocumentBuilderSetFormatting();
    DocumentBuilderMovingCursor();
    InsertDoc();
    DocumentBuilderInsertTCFieldsAtText();
    RemoveTOCFromDocument();
    SetCompatibilityOptions();
    Setuplanguagepreferences();
    WorkingWithRevisions();
    ExtractContentUsingField();
    ExtractTextOnly();
    DocumentBuilderInsertTOC();
    DocumentBuilderInsertTCField();
    CheckBoxTypeContentControl();
    RichTextBoxContentControl();
    ComboBoxContentControl();
    UpdateContentControls();
    CleansUnusedStylesandLists();
    DocumentPageSetup();

    // Fields
    // =====================================================
    RemoveField();
    ConvertFieldsInDocument();
    ConvertFieldsInBody();
    ConvertFieldsInParagraph();
    SpecifylocaleAtFieldlevel();
    InsertFormFields();
    FormFieldsGetFormFieldsCollection();
    FormFieldsGetByName();
    FormFieldsWorkWithProperties();
    RenameMergeFields();
    InsertNestedFields();
    ChangeLocale();
    UpdateDocFields();
    UseOfficeMathProperties();
    InsertField();
    InsertMergeFieldUsingDOM();
    InsertMailMergeAddressBlockFieldUsingDOM();
    InsertAdvanceFieldWithoutDocumentBuilder();
    InsertASKFieldWithoutDocumentBuilder();
    InsertAuthorField();
    // TODO (std_string) : investigate why this example doesn't work properly
    //ChangeFieldUpdateCultureSource();
    GetFieldNames();
    InsertTOAFieldWithoutDocumentBuilder();
    InsertIncludeTextFieldWithoutDocumentBuilder();
    EvaluateIFCondition();

    // Images
    // =====================================================
    AddWatermark();
    RemoveWatermark();
    CompressImages();
    ExtractImagesToFiles();
    InsertBarcodeImage();

    // Ranges
    // =====================================================
    RangesDeleteText();
    RangesGetText();

    // Theme
    // =====================================================
    ManipulateThemeProperties();

    // Node
    // =====================================================
    ExNode();

    // Hyperlink
    // =====================================================
    ReplaceHyperlinks();


    // Styles
    // =====================================================
    ChangeStyleOfTOCLevel();
    ChangeTOCTabStops();
    ExtractContentBasedOnStyles();
    InsertStyleSeparator();

    // Tables
    // =====================================================
    AddRemoveColumn();
    ApplyFormatting();
    ApplyStyle();
    AutoFitTableToContents();
    AutoFitTableToFixedColumnWidths();
    AutoFitTableToWindow();
    CloneTable();
    ExtractOrReplaceText();
    FindingIndex();
    InsertTableDirectly();
    InsertTableUsingDocumentBuilder();
    JoiningAndSplittingTable();
    KeepTablesAndRowsBreaking();
    MergedCells();
    RepeatRowsOnSubsequentPages();
    SpecifyHeightAndWidth();
    TablePosition();

    // Sections
    // =====================================================
    AddDeleteSection();
    AppendSectionContent();
    CloneSection();
    CopySection();
    DeleteHeaderFooterContent();
    DeleteSectionContent();
    SectionsAccessByIndex();

    // StructuredDocumentTag 
    // =====================================================
    WorkingWithSDT();

    // EndNote and Footnote
    // =====================================================
    WorkingWithFootnote();

    // List
    // =====================================================
    WorkingWithList();

    // =====================================================
    // =====================================================
    // Mail-Merge
    // =====================================================
    // =====================================================
    SimpleMailMerge();
    MailMergeFormFields();
    ExecuteArray();
    NestedMailMergeCustom();
    HandleMailMergeSwitches();
    MailMergeAndConditionalField();

    std::cout << "=====================================================" << std::endl << "Examples Finished." << std::endl;

    return 0;
}

