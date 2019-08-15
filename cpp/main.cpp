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
    PageSplitter();
    ConvertDocumentToEPUB();
    ConvertDocumentToHtmlWithRoundtrip();
    Doc2Pdf();
    ExportFontsAsBase64();
    ExportResourcesUsingHtmlSaveOptions();
    LoadAndSaveHtmlFormFieldAsContentControlInDOCX();
    SaveDocWithHtmlSaveOptions();
    SpecifySaveOption();
    SplitIntoHtmlPages();
    WorkingWithOoxml();
    WorkingWithDoc();

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
    ConvertNumPageFields();
    UpdatePageLayout();

    // Find and Replace
    // =====================================================
    FindAndHighlight();
    ReplaceWithString();
    ReplaceWithRegex();
    ReplaceWithEvaluator();
    FindReplaceUsingMetaCharacters();
    ReplaceTextWithField();
    ReplaceHtmlTextWithMetaCharacters();
    ReplaceWithHTML();

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
    RemoveRegionText();

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
    WorkingWithSaveOptions();
    DocumentBuilderInsertHorizontalRule();
    PageNumbersOfNodes();
    CheckDMLTextEffect();
    ParagraphStyleSeparator();
    WorkingWithMarkdownFeatures();
    CompareDocument();

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
    ChangeFieldUpdateCultureSource();
    GetFieldNames();
    InsertTOAFieldWithoutDocumentBuilder();
    InsertIncludeTextFieldWithoutDocumentBuilder();
    EvaluateIFCondition();
    FieldUpdateCulture();
    FormatFieldResult();
    InsertFieldNone();

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

    // Charts
    // =====================================================
    ChartNumberFormat();
    CreateChartUsingShape();
    CreateColumnChart();
    InsertAreaChart();
    InsertBubbleChart();
    InsertScatterChart();
    WorkWithChartDataLabel();
    WorkWithSingleChartDataPoint();
    WorkWithSingleChartSeries();
    WorkingWithChartAxis();

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
    CopyStyles();

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
    InsertTableFromHtml();

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

    // Linked TextBoxes
    // =====================================================
    WorkingWithLinkedTextboxes();

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
    MailMergeCleanUp();

    // =====================================================
    // =====================================================
    // Rendering and Printing
    // =====================================================
    // =====================================================
    DocumentLayoutHelper();
    EmbeddedFontsInPDF();
    EmbeddingWindowsStandardFonts();
    EnumerateLayoutElements();
    HyphenateWordsOfLanguages();
    LoadHyphenationDictionaryForLanguage();
    ReceiveNotificationsOfFont();
    RenderShape();
    SetFontSettings();
    SetFontsFoldersMultipleFolders();
    SetFontsFoldersSystemAndCustomFolder();
    SetHorizontalAndVerticalImageResolution();
    SetTrueTypeFontsFolder();
    SpecifyDefaultFontWhenRendering();
    WorkingWithFontSources();
    WorkingWithPdfSaveOptions();

    std::cout << "=====================================================" << std::endl << "Examples Finished." << std::endl;

    return 0;
}

