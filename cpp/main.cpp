#include "stdafx.h"
#include <iostream>

#if defined(__GNUC__)
#include <chrono>
#include <thread>

#include <Aspose.Words.Cpp/AsposeWordsLibrary.h>
#endif

#include "examples.h"

int main()
{
#if defined(__GNUC__)
	// Workaround for issue with pthread https://gcc.gnu.org/bugzilla/show_bug.cgi?id=60662
	std::this_thread::sleep_for(std::chrono::milliseconds{1});
#endif

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
    SpecifyMarkdownSaveOptions();
	WorkingWithVbaReferenceCollection();
    ConvertWordDocument();

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
    WorkingWithImportFormatOptions();
    WorkingWithMarkdownFeatures();
    WorkingWithRevisions();
	WorkingWithRevisionOptions();
    WorkingWithRtfSaveOptions();
    WorkingWithSaveOptions();
    WorkingWithWebExtension();
    WorkWithCleanupOptions();
    WorkWithWatermark();
    ShowGrammaticalAndSpellingErrors();
    SplitDocument();

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
#if 0
    // Input document are missing
	FormFieldsFontFormatting();
#endif
    GetFieldNames();
    InsertAdvanceFieldWithoutDocumentBuilder();
    InsertASKFieldWithoutDocumentBuilder();
    InsertAuthorField();
    InsertFields();
    InsertFieldNone();
    InsertFormFields();

#if !defined(__GNUC__) || defined(NDEBUG) // WORDSCPP-1010 assertion failure
    InsertIncludeFieldWithoutDocumentBuilder();
#endif
    InsertMailMergeAddressBlockFieldUsingDOM();
    InsertMergeFieldUsingDOM();
    InsertNestedFields();
    InsertTOAFieldWithoutDocumentBuilder();
    RemoveField();
    RenameMergeFields();
    SpecifylocaleAtFieldlevel();
    UpdateDocFields();
    UseOfficeMathProperties();
    FieldDisplayResults();

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
    ReplaceTextInTable();

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
    AppendWithImportFormatOptions(); 
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
    ModifyPageSetupInAllSectionsOfDocument(); 
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
#if !defined(__GNUC__) // WORDSCPP-1011 throw std::out_of_range without required fonts
    DocumentLayoutHelper();
#endif

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
#endif
    ReadActiveXControlProperties();
    ReceiveNotificationsOfFont();
    RenderShape();
    ResourceSteamFontSource();
    SaveAsMultipageTiff();
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

#if defined(__GNUC__)
	// Unload thread created by Aspose.Words for C++
	Aspose::Words::AsposeWordsLibrary::PrepareForUnload();
#endif 

    return 0;
}

