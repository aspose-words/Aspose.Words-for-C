#include <iostream>

#include "examples.h"

int main()
{
    std::cout << "Open main.cpp \nIn main() function uncomment the example that you want to run.\n"
              << "=====================================================\n";

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

    // =====================================================
    // =====================================================
    // Loading and Saving
    // =====================================================
    // =====================================================
    LoadAndSaveToDisk();
    LoadAndSaveToStream();
    CreateDocument();
    ConvertDocumentToByte();

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

    // Bookmarks
    // =====================================================
    CopyBookmarkedText();
    UntangleRowBookmarks();
    BookmarkTable();
    BookmarkNameAndText();
    AccessBookmarks();
    CreateBookmark();

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

    // Fields
    // =====================================================
    RemoveField();
    SpecifylocaleAtFieldlevel();
    InsertFormFields();
    FormFieldsGetFormFieldsCollection();
    FormFieldsGetByName();
    FormFieldsWorkWithProperties();
    RenameMergeFields();

    // Images
    // =====================================================
    AddWatermark();
    RemoveWatermark();
    CompressImages();
    ExtractImagesToFiles();

    // Ranges
    // =====================================================
    RangesDeleteText();
    RangesGetText();

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

    // EndNote and Footnote
    WorkingWithFootnote();

    return 0;
}

