#include "stdafx.h"
#include <iostream>

#include "examples.h"

int main()
{
    std::cout << "Examples:" << std::endl << "=====================================================" << std::endl << std::endl;

    // =====================================================
    // =====================================================
    // Quick Start
    // =====================================================
    // =====================================================
    AppendDocuments();
    ApplyLicense();
    ApplyLicenseFromFile();
    ApplyLicenseFromStream();
#if 0 // Metered license isn't supported
    ApplyMeteredLicense(); 
#endif 
    FindAndReplace();
    HelloWorld();
    UpdateFields();
    WorkingWithNodes();


    // =====================================================
    // =====================================================
    // Loading and Saving
    // =====================================================
    // =====================================================
    AccessAndVerifySignature();
    CheckFormat();
    ConvertDocumentToByte();
    ConvertDocumentToEPUB();
    ConvertDocumentToHTML();
    ConvertDocumentToPCL();
    CreateDocument();
    DetectDocumentSignatures();
    DigitallySignedPdfUsingCertificateHolder();
    Doc2Pdf();
    ImageToPdf(); 
    Load_Options();
    LoadAndSaveHtmlFormFieldAsContentControlInDOCX();
    LoadAndSaveToDisk();
    LoadAndSaveToStream();
    LoadTxt();
    OpenEncryptedDocument();
    PageSplitter();
    SaveDocWithHtmlSaveOptions();
    SaveOptionsHtmlFixed();
    SpecifySaveOption();
    SplitIntoHtmlPages();
    WorkingWithDoc();
    WorkingWithOoxml();
    WorkingWithRTF();
    WorkingWithTxt();
    WorkingWithVbaMacros(); 

    // =====================================================
    // =====================================================
    // Mail-Merge
    // =====================================================
    // =====================================================
#if 0
    // DataSet isn't supported
    ApplyCustomLogicToEmptyRegions(); 
    MailMergeUsingMustacheSyntax();
    RemoveEmptyRegions(); 
    RemoveRowsFromTable(); 
    XMLMailMerge(); 
    NestedMailMerge(); 

    // OleDbXXX & DataTable aren't supported
    ExecuteWithRegionsDataTable(); 
    MailMergeAlternatingRows(); 
    MailMergeImageFromBlob();
    MultipleDocsInMailMerge();
    ProduceMultipleDocuments(); 

    // Linq2Xml isn't supported
    LINQtoXMLMailMerge();
#endif 
    ExecuteArray();
    HandleMailMergeSwitches();
    MailMergeAndConditionalField();
    MailMergeCleanUp();
    MailMergeFormFields();
    NestedMailMergeCustom();
    SimpleMailMerge();
    MailMergeImageField();

    // =====================================================
    // =====================================================
    // Programming with Documents
    // =====================================================
    // =====================================================

    // Bookmarks
    // =====================================================
    AccessBookmarks();
    BookmarkNameAndText();
    BookmarkTable();
    CopyBookmarkedText();
    CreateBookmark();
    UntangleRowBookmarks();
    ShowHideBookmarks();

    // Charts
    // =====================================================
    ChartNumberFormat();
    CreateChartUsingShape();
    CreateColumnChart();
    InsertAreaChart();
    InsertBubbleChart();
    InsertScatterChart();
    WorkingWithChartAxis();
    WorkWithChartDataLabels();
    WorkWithSingleChartDataPoint();
    WorkWithSingleChartSeries();

    // Comments
    // =====================================================
    AddComments();
    AnchorComment();
    CommentReply();
    ProcessComments();
    RemoveRegionText();

    // ConvertUtil
    // =====================================================
    UtilityClasses();

    // Document
    // =====================================================
    AccessStyles();
    AddGroupShapeToDocument();
    CheckBoxTypeContentControl();
    CheckDMLTextEffect();
    CleansUnusedStylesandLists();
    CloningDocument();
    ComboBoxContentControl();
    CompareDocument();
    CreateHeaderFooterUsingDocBuilder();
    DocProperties(); 
    DocumentBuilderBuildTable();
    DocumentBuilderHorizontalRule();
    DocumentBuilderInsertBookmark();
    DocumentBuilderInsertBreak();
    DocumentBuilderInsertElements();
    DocumentBuilderInsertImage();
    DocumentBuilderInsertParagraph();
    DocumentBuilderInsertTCField();
    DocumentBuilderInsertTCFieldsAtText();
    DocumentBuilderInsertTOC();
    DocumentBuilderMovingCursor();
    DocumentBuilderSetFormatting();
    DocumentPageSetup();
    ExtractContentBetweenBlockLevelNodes();
    ExtractContentBetweenBookmark();
    ExtractContentBetweenCommentRange();
    ExtractContentBetweenParagraphs();
    ExtractContentBetweenParagraphStyles();
    ExtractContentBetweenRuns();
    ExtractContentUsingDocumentVisitor();
    ExtractContentUsingField();
    ExtractTableOfContents();
    ExtractTextOnly();
#if 0
    // Depends on Aspose.BarCode
    GenerateACustomBarCodeImage();
#endif
    GetFontLineSpacing();
    GetVariables();
    InsertDoc();
    PageNumbersOfNodes();
    ParagraphStyleSeparator();
    ProtectDocument();
    RemoveBreaks();
    RemoveFooters();
    RemoveTOCFromDocument();
    RichTextBoxContentControl();
    SetCompatibilityOptions();
    Setuplanguagepreferences();
    SetViewOption();
    SigningSignatureLine();
    UpdateContentControls();
    WriteAndFont();
#if 0
    // Source document is missing
    WorkingWithImportFormatOptions();
#endif
    WorkingWithMarkdownFeatures();
    WorkingWithRevisions();
    WorkingWithRtfSaveOptions();
    WorkingWithSaveOptions();
    WorkingWithWebExtension();
    WorkWithCleanupOptions();
    WorkWithWatermark();
    ShowGrammaticalAndSpellingErrors();

    // EndNote and Footnote
    // =====================================================
    WorkingWithFootnote();

    // Fields
    // =====================================================
    ChangeFieldUpdateCultureSource();
    ChangeLocale();
    ConvertFieldsInBody();
    ConvertFieldsInDocument();
    ConvertFieldsInParagraph();
    EvaluateIFCondition();
    FieldUpdateCulture();
    FormatFieldResult();
    FormFieldsGetByName();
    FormFieldsGetFormFieldsCollection();
    FormFieldsWorkWithProperties();
    GetFieldNames();
    InsertAdvanceFieldWithoutDocumentBuilder();
    InsertASKFieldWithoutDocumentBuilder();
    InsertAuthorField();
    InsertFields();
    InsertFieldNone();
    InsertFormFields();
    InsertIncludeFieldWithoutDocumentBuilder();
    InsertMailMergeAddressBlockFieldUsingDOM();
    InsertMergeFieldUsingDOM();
    InsertNestedFields();
    InsertTOAFieldWithoutDocumentBuilder();
    RemoveField();
    RenameMergeFields();
    SpecifylocaleAtFieldlevel();
    UpdateDocFields();
    UseOfficeMathProperties();

    // Find and Replace
    // =====================================================
    FindAndHighlight();
    FindReplaceUsingMetaCharacters();
    ReplaceHtmlTextWithMetaCharacters();
    ReplaceTextWithField();
    ReplaceWithEvaluator();
    ReplaceWithHTML();
    ReplaceWithRegex();
    ReplaceWithString();
    UsingLegacyOrder();
    IgnoreText();

    // Hyperlink
    // =====================================================
    ReplaceHyperlinks();

    // Images
    // =====================================================
    AddImageToEachPage();
    AddWatermark();
    CompressImages();
    ExtractImagesToFiles();
    InsertBarcodeImage();
    RemoveWatermark();

    // Joining and Appending
    // =====================================================
    AppendDocumentManually();
#if 0
    // Source document is missing
    AppendWithImportFormatOptions(); 
#endif
    BaseDocument();
    ConvertNumPageFields();
    DifferentPageSetup();
    JoinContinuous();
    JoinNewPage();
    KeepSourceFormatting();
    KeepSourceTogether();
    LinkHeadersFooters();
    ListKeepSourceFormatting();
    ListUseDestinationStyles();
    PrependDocument();
    RemoveSourceHeadersFooters();
    RestartPageNumbering();
    SimpleAppendDocument();
    UnlinkHeadersFooters();
    UpdatePageLayout();
    UseDestinationStyles();

    // Linked TextBoxes
    // =====================================================
    WorkingWithLinkedTextboxes();

    // List
    // =====================================================
    WorkingWithList();

    // Node
    // =====================================================
    ExNode();

    // Ranges
    // =====================================================
    RangesDeleteText();
    RangesGetText();

    // Sections
    // =====================================================
    AddDeleteSection();
    AppendSectionContent();
    CloneSection();
    CopySection();
    DeleteHeaderFooterContent();
    DeleteSectionContent();
#if 0
    // Source document is missing
    ModifyPageSetupInAllSectionsOfDocument(); 
#endif
    SectionsAccessByIndex();

    // Shapes
    // =====================================================
    WorkingWithShapes(); 

    // StructuredDocumentTag 
    // =====================================================
    WorkingWithSDT();

    // Styles
    // =====================================================
    ChangeStyleOfTOCLevel();
    ChangeTOCTabStops();
    CopyStyles();
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
#if 0
    // DataTable isn't supported
    BuildTableFromDataTable(); 
#endif
    CloneTable();
    ExtractOrReplaceText();
    FindingIndex();
    InsertTableDirectly();
    InsertTableFromHtml();
    InsertTableUsingDocumentBuilder();
    JoiningAndSplittingTable();
    KeepTablesAndRowsBreaking();
    MergedCells(); 
    RepeatRowsOnSubsequentPages();
    SpecifyHeightAndWidth();
    TablePosition();

    // Theme
    // =====================================================
    ManipulateThemeProperties();

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
    ImageColorFilters(); 
    LoadHyphenationDictionaryForLanguage();
#if 0
    // GUI isn't supported
    PrintProgressDialog();
    Print_CachePrinterSettings;
    // OleControl isn't supported
    ReadActiveXControlProperties();
#endif
    ReceiveNotificationsOfFont();
    RenderShape();
    ResourceSteamFontSource();
#if 0
    // Tiff support is limited and unstable
    SaveAsMultipageTiff();
#endif
    WorkingWithFontSettings();
    SetFontsFoldersMultipleFolders();
    SetFontsFoldersSystemAndCustomFolder();
    SetHorizontalAndVerticalImageResolution();
    SetTrueTypeFontsFolder();
    SpecifyDefaultFontWhenRendering();
#if 0
    // Source document is missing
    WorkingWithFontResolution(); 
#endif
    WorkingWithFontSources();
    WorkingWithPdfSaveOptions();
    SetFontsFolders();
    SetFontsFoldersDefaultInstance();
    SetFontsFoldersWithPriority();

    std::cout << "=====================================================" << std::endl << "Examples Finished." << std::endl;

    return 0;
}

