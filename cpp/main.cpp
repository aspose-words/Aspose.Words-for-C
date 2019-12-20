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
    //AccessAndVerifySignature(); // Cryptography isn't supported
    CheckFormat();
    ConvertDocumentToByte();
    ConvertDocumentToEPUB();
    ConvertDocumentToHtmlWithRoundtrip();
    ConvertDocumentToPCL();
    CreateDocument();
    DetectDocumentSignatures();
    //DigitallySignedPdf(); // Cryptography isn't supported
    //DigitallySignedPdfUsingCertificateHolder(); // Cryptography isn't supported
    Doc2Pdf();
    ExportFontsAsBase64();
    ExportResourcesUsingHtmlSaveOptions();
    //ImageToPdf(); // System::Drawing::Image::get_FrameDimensionsList() isn't implemented in the asposecpp lib
    Load_Options(); // Cryptography isn't supported
    LoadAndSaveHtmlFormFieldAsContentControlInDOCX();
    LoadAndSaveToDisk();
    LoadAndSaveToStream();
    LoadTxt();
    //OpenEncryptedDocument(); // Cryptography isn't supported
    PageSplitter();
    SaveDocWithHtmlSaveOptions();
    SaveOptionsHtmlFixed();
    SpecifySaveOption();
    SplitIntoHtmlPages();
    WorkingWithDoc(); // Cryptography isn't supported + Source document is missing
    WorkingWithOoxml(); // Cryptography isn't supported
    WorkingWithRTF();
    WorkingWithTxt(); // Source document is missing
    WorkingWithVbaMacros();

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
    //SigningSignatureLine(); // Cryptography isn't supported
    UpdateContentControls();
    //WorkingWithImportFormatOptions(); // Source document is missing
    WorkingWithMarkdownFeatures();
    WorkingWithRevisions();
    WorkingWithSaveOptions();
    WorkingWithWebExtension();
    WriteAndFont();

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
    InsertField();
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
    MergedCells(); // error in test
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
    SetFontSettings();
    SetFontsFoldersMultipleFolders();
    SetFontsFoldersSystemAndCustomFolder();
    SetHorizontalAndVerticalImageResolution();
    SetTrueTypeFontsFolder();
    SpecifyDefaultFontWhenRendering();
    //WorkingWithFontResolution(); // Source document is missing
    WorkingWithFontSources();
    WorkingWithPdfSaveOptions();

    std::cout << "=====================================================" << std::endl << "Examples Finished." << std::endl;

    return 0;
}

