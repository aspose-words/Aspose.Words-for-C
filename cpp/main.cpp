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
    ApplyLicense();
    ApplyLicenseFromFile();
    ApplyLicenseFromStream();
    // TODO(std_string) : Metered license isn't supported
    //ApplyMeteredLicense();
    AppendDocuments();
    FindAndReplace();
    HelloWorld();
    UpdateFields();
    WorkingWithNodes();

    // =====================================================
    // =====================================================
    // Loading and Saving
    // =====================================================
    // =====================================================
    // TODO (std_string) : Cryptography isn't supported
    //AccessAndVerifySignature();
    CheckFormat();
    ConvertDocumentToByte();
    ConvertDocumentToEPUB();
    ConvertDocumentToHtmlWithRoundtrip();
    ConvertDocumentToPCL();
    CreateDocument();
    DetectDocumentSignatures();
    // TODO (std_string) : Cryptography isn't supported 
    //DigitallySignedPdf();
    // TODO (std_string) : Cryptography isn't supported
    //DigitallySignedPdfUsingCertificateHolder();
    Doc2Pdf();
    ExportFontsAsBase64();
    ExportResourcesUsingHtmlSaveOptions();
    // TODO (std_string) : System::Drawing::Image::get_FrameDimensionsList() isn't implemented in the asposecpp lib
    //ImageToPdf();
    LoadAndSaveHtmlFormFieldAsContentControlInDOCX();
    LoadAndSaveToDisk();
    LoadAndSaveToStream();
    Load_Options();
    LoadTxt();
    // TODO (std_string) : Cryptography isn't supported
    //OpenEncryptedDocument();
    PageSplitter();
    SaveDocWithHtmlSaveOptions();
    SaveOptionsHtmlFixed();
    SpecifySaveOption();
    SplitIntoHtmlPages();
    // TODO(std_string) : Cryptography isn't supported
    WorkingWithDoc();
    WorkingWithOoxml();
    WorkingWithRTF();
    WorkingWithTxt();
    // TODO Input file are missing
	// WorkingWithVbaMacros();
    // =====================================================
    // =====================================================
    // Programming with Documents
    // =====================================================
    // =====================================================

    // Joining and Appending
    // =====================================================
    AppendDocumentManually();
    // TODO(std_string) : The source document is missing
    // AppendWithImportFormatOptions();
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

    // Bookmarks
    // =====================================================
    AccessBookmarks();
    BookmarkNameAndText();
    BookmarkTable();
    CopyBookmarkedText();
    CreateBookmark();
    UntangleRowBookmarks();

    // Shapes
    // =====================================================
    WorkingWithShapes();

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
    // DocProperties();
    DocumentBuilderBuildTable();
    DocumentBuilderInsertBookmark();
    DocumentBuilderInsertBreak();
    DocumentBuilderInsertElements();
    DocumentBuilderInsertHorizontalRule();
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
    // TODO (std_string) : Dependency from Aspose.BarCode 
    //GenerateACustomBarCodeImage.Run();
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
    // TODO (std_string) : Cryptography isn't supported
    //SigningSignatureLine();
    UpdateContentControls();
    // TODO(std_string) : The source document is missing 
    //WorkingWithImportFormatOptions();
    WorkingWithMarkdownFeatures();
    WorkingWithRevisions();
    WorkingWithSaveOptions();
    WriteAndFont();

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
    InsertIncludeTextFieldWithoutDocumentBuilder();
    InsertMailMergeAddressBlockFieldUsingDOM();
    InsertMergeFieldUsingDOM();
    InsertNestedFields();
    InsertTOAFieldWithoutDocumentBuilder();
    RemoveField();
    RenameMergeFields();
    SpecifylocaleAtFieldlevel();
    UpdateDocFields();
    UseOfficeMathProperties();

    // Images
    // =====================================================
    AddImageToEachPage();
    AddWatermark();
    CompressImages();
    ExtractImagesToFiles();
    InsertBarcodeImage();
    RemoveWatermark();

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
    WorkingWithChartAxis();
    WorkWithChartDataLabel();
    WorkWithSingleChartDataPoint();
    WorkWithSingleChartSeries();

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
    // TODO(std_string) : isn't compiled due to using of DataTable
    //BuildTableFromDataTable();
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

    // Sections
    // =====================================================
    AddDeleteSection();
    AppendSectionContent();
    CloneSection();
    CopySection();
    DeleteHeaderFooterContent();
    DeleteSectionContent();
    // TODO(std_string) : The source document is missing
    //ModifyPageSetupInAllSectionsOfDocument();
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
    // TODO(std_string) : isn't compiled due to using of DataSet
    //ApplyCustomLogicToEmptyRegions();
    ExecuteArray();
    // TODO(std_string) : isn't compiled due to using of OleDbXXX & DataTable
    //ExecuteWithRegionsDataTable();
    HandleMailMergeSwitches();
    // TODO(std_string) : isn't compiled due to using of Linq2Xml
    //LINQtoXMLMailMerge();
    // TODO(std_string) : isn't compiled due to using of DataTable
    //MailMergeAlternatingRows();
    MailMergeAndConditionalField();
    MailMergeCleanUp();
    MailMergeFormFields();
    // TODO(std_string) : isn't compiled due to using of OleDbXXX & IDataReader
    //MailMergeImageFromBlob();
    // TODO(std_string) : isn't compiled due to using of DataSet
    //MailMergeUsingMustacheSyntax.Run();
    // TODO(std_string) : isn't compiled due to using of OleDbXXX & DataTable
    //MultipleDocsInMailMerge();
    // TODO(std_string) : isn't compiled due to using of DataSet
    //NestedMailMerge();
    NestedMailMergeCustom();
    // TODO(std_string) : OleDbXXX & DataTable
    //ProduceMultipleDocuments();
    // TODO(std_string) : isn't compiled due to using of DataSet
    //RemoveEmptyRegions();
    // TODO(std_string) : isn't compiled due to using of DataSet
    //RemoveRowsFromTable();
    SimpleMailMerge();
    // TODO(std_string) : isn't compiled due to using of DataSet
    //XMLMailMerge();

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
    // TODO (std_string) : saving into images don't work properly
    //ImageColorFilters();
    LoadHyphenationDictionaryForLanguage();
    // TODO (std_string) : doesn't work due to using of GUI
    //PrintProgressDialog();
    // TODO (std_string) : doesn't work due to using of GUI
    //Print_CachePrinterSettings;
    // TODO (std_string) : doesn't work due to using of OleControl
    //ReadActiveXControlProperties();
    ReceiveNotificationsOfFont();
    RenderShape();
    ResourceSteamFontSource();
    // TODO (std_string) : saving into images don't work properly
    //SaveAsMultipageTiff();
    SetFontSettings();
    SetFontsFoldersMultipleFolders();
    SetFontsFoldersSystemAndCustomFolder();
    SetHorizontalAndVerticalImageResolution();
    SetTrueTypeFontsFolder();
    SpecifyDefaultFontWhenRendering();
    // TODO (std_string) : The source document is missing
    //WorkingWithFontResolution();
    WorkingWithFontSources();
    WorkingWithPdfSaveOptions();

    std::cout << "=====================================================" << std::endl << "Examples Finished." << std::endl;

    return 0;
}

