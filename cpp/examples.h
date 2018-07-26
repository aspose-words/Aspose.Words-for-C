#pragma once

#include "RunExamples.h"

// Quick Start

void AppendDocuments();
void ApplyLicense();
void FindAndReplace();
void HelloWorld();
void WorkingWithNodes();

// Loading and Saving

void LoadAndSaveToDisk();
void LoadAndSaveToStream();
void CreateDocument();
void CheckFormat();
void ImageToDoc();
void ConvertDocumentToByte();

// =================================
// Programming with Documents
// =================================

// Joining and Appending

void SimpleAppendDocument();
void KeepSourceFormatting();
void UseDestinationStyles();
void JoinContinuous();
void JoinNewPage();
void RestartPageNumbering();
void LinkHeadersFooters();
void UnlinkHeadersFooters();
void RemoveSourceHeadersFooters();
void DifferentPageSetup();
void ListUseDestinationStyles();
void ListKeepSourceFormatting();
void KeepSourceTogether();
void BaseDocument();
void AppendDocumentManually();
void PrependDocument();

// Find and Replace

void FindAndHighlight();
void ReplaceWithString();
void ReplaceWithRegex();
void ReplaceWithEvaluator();
void FindReplaceUsingMetaCharacters();

// Bookmarks

void CopyBookmarkedText();
void UntangleRowBookmarks();
void BookmarkTable();
void BookmarkNameAndText();
void AccessBookmarks();
void CreateBookmark();

// Comments
void ProcessComments();
void AddComments();
void AnchorComment();
void CommentReply();

// ConvertUtil
void UtilityClasses();

// Document
void ExtractContentBetweenParagraphs();
void ExtractContentBetweenBlockLevelNodes();
void ExtractContentBetweenParagraphStyles();
void ExtractContentBetweenRuns();
void ExtractContentBetweenBookmark();
void ExtractContentBetweenCommentRange();
void RemoveBreaks();
void CloningDocument();
void ProtectDocument();
void AccessStyles();
void GetVariables();
void SetViewOption();
void CreateHeaderFooterUsingDocBuilder();
void RemoveFooters();
void AddGroupShapeToDocument();
void DocProperties();
void WriteAndFont();
void DocumentBuilderInsertParagraph();
void DocumentBuilderBuildTable();
void DocumentBuilderInsertBreak();
void DocumentBuilderInsertImage();
void DocumentBuilderInsertBookmark();
void DocumentBuilderInsertElements();
void DocumentBuilderSetFormatting();
void DocumentBuilderMovingCursor();
void InsertDoc();
void DocumentBuilderInsertTCFieldsAtText();
void RemoveTOCFromDocument();
void SetCompatibilityOptions();
void Setuplanguagepreferences();

// Fields
void RemoveField();
void ConvertFieldsInDocument();
void ConvertFieldsInBody();
void ConvertFieldsInParagraph();
void SpecifylocaleAtFieldlevel();
void InsertFormFields();
void FormFieldsGetFormFieldsCollection();
void FormFieldsGetByName();
void FormFieldsWorkWithProperties();
void RenameMergeFields();

// Images
void AddWatermark();
void RemoveWatermark();
void CompressImages();
void ExtractImagesToFiles();

// Ranges
void RangesGetText();
void RangesDeleteText();

// Charts

// Node
void ExNode();

// Hyperlink
void ReplaceHyperlinks();

// Styles
void ExtractContentBasedOnStyles();
void ChangeStyleOfTOCLevel();
void ChangeTOCTabStops();
void InsertStyleSeparator();

// Tables
void AddRemoveColumn();
void ApplyFormatting();
void ApplyStyle();
void AutoFitTableToContents();
void AutoFitTableToFixedColumnWidths();
void AutoFitTableToWindow();
void CloneTable();
void ExtractOrReplaceText();
void FindingIndex();
void InsertTableDirectly();
void InsertTableUsingDocumentBuilder();
void JoiningAndSplittingTable();
void KeepTablesAndRowsBreaking();
void MergedCells();
void RepeatRowsOnSubsequentPages();
void SpecifyHeightAndWidth();
void TablePosition();

// Sections
void AddDeleteSection();
void AppendSectionContent();
void CloneSection();
void CopySection();
void DeleteHeaderFooterContent();
void DeleteSectionContent();
void SectionsAccessByIndex();

// EndNote and Footnote
void WorkingWithFootnote();
