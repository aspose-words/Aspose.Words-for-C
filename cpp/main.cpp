#include "stdafx.h"
#include <iostream>

#include "examples.h"

int main()
{
    std::cout << "Examples:" << std::endl << "=====================================================" << std::endl << std::endl;

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
    //ImageToPdf(); // System::Drawing::Image::get_FrameDimensionsList() isn't implemented in the asposecpp lib
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
    WorkingWithDoc(); // Source document is missing
    WorkingWithOoxml();
    WorkingWithRTF();
    WorkingWithTxt(); // Source document is missing
    WorkingWithVbaMacros(); // raised NullReferenceException due to bad source document

    // =====================================================
    // =====================================================
    // Mail-Merge
    // =====================================================
    // =====================================================
    //ApplyCustomLogicToEmptyRegions(); // isn't compiled due to using of DataSet
    ExecuteArray();
    //ExecuteWithRegionsDataTable(); // isn't compiled due to using of OleDbXXX & DataTable
    HandleMailMergeSwitches();
    //LINQtoXMLMailMerge(); // isn't compiled due to using of Linq2Xml
    //MailMergeAlternatingRows(); // isn't compiled due to using of DataTable
    MailMergeAndConditionalField();
    MailMergeCleanUp();
    MailMergeFormFields();
    //MailMergeImageFromBlob(); // isn't compiled due to using of OleDbXXX & IDataReader
    //MailMergeUsingMustacheSyntax.Run(); // isn't compiled due to using of DataSet
    //MultipleDocsInMailMerge(); // isn't compiled due to using of OleDbXXX & DataTable
    //NestedMailMerge(); // isn't compiled due to using of DataSet
    NestedMailMergeCustom();
    //ProduceMultipleDocuments(); // isn't compiled due to using of OleDbXXX & DataTable
    //RemoveEmptyRegions(); // isn't compiled due to using of DataSet
    //RemoveRowsFromTable(); // isn't compiled due to using of DataSet
    SimpleMailMerge();
    //XMLMailMerge(); // isn't compiled due to using of DataSet
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
    BookmarkTable(); // Source document is missing
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
    DocProperties(); // exception at execution
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
    //GenerateACustomBarCodeImage.Run(); // Depend from Aspose.BarCode
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
    //WorkingWithImportFormatOptions(); // Source document is missing
    WorkingWithMarkdownFeatures();
    WorkingWithRevisions();
    WorkingWithRtfSaveOptions();
    WorkingWithSaveOptions();
    WorkingWithWebExtension();

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
    // AppendWithImportFormatOptions(); // Source document is missing
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
    //ModifyPageSetupInAllSectionsOfDocument(); // Source document is missing
    SectionsAccessByIndex();

    // Shapes
    // =====================================================
    WorkingWithShapes(); // Source document is missing

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
    //BuildTableFromDataTable(); // isn't compiled due to using of DataTable
    CloneTable();
    ExtractOrReplaceText();
    FindingIndex();
    InsertTableDirectly();
    InsertTableFromHtml();
    InsertTableUsingDocumentBuilder();
    JoiningAndSplittingTable();
    KeepTablesAndRowsBreaking();
    MergedCells(); // raised NullReferenceException
    RepeatRowsOnSubsequentPages();
    SpecifyHeightAndWidth();
    TablePosition();

    // Theme
    // =====================================================
    ManipulateThemeProperties();

    // =====================================================
    // =====================================================
    // Quick Start
    // =====================================================
    // =====================================================
    AppendDocuments();
    ApplyLicense();
    ApplyLicenseFromFile();
    ApplyLicenseFromStream();
    //ApplyMeteredLicense(); // Metered license isn't supported
    FindAndReplace();
    HelloWorld();
    UpdateFields();
    WorkingWithNodes();

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
    ImageColorFilters(); // saving into images don't work properly
    LoadHyphenationDictionaryForLanguage();
    //PrintProgressDialog(); // using of GUI
    //Print_CachePrinterSettings; // using of GUI
    //ReadActiveXControlProperties(); // doesn't work due to using of OleControl
    ReceiveNotificationsOfFont();
    RenderShape();
    ResourceSteamFontSource();
    //SaveAsMultipageTiff(); // saving into images don't work properly
    WorkingWithFontSettings();
    SetFontsFoldersMultipleFolders();
    SetFontsFoldersSystemAndCustomFolder();
    SetHorizontalAndVerticalImageResolution();
    SetTrueTypeFontsFolder();
    SpecifyDefaultFontWhenRendering();
    //WorkingWithFontResolution(); // Source document is missing
    WorkingWithFontSources();
    WorkingWithPdfSaveOptions();
    SetFontsFolders();
    SetFontsFoldersDefaultInstance();
    SetFontsFoldersWithPriority();

    std::cout << "=====================================================" << std::endl << "Examples Finished." << std::endl;

    return 0;
}

