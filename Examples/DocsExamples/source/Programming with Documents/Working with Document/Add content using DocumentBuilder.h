#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Bookmark.h>
#include <Aspose.Words.Cpp/BookmarkCollection.h>
#include <Aspose.Words.Cpp/BookmarkEnd.h>
#include <Aspose.Words.Cpp/BookmarkStart.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/HorizontalRuleAlignment.h>
#include <Aspose.Words.Cpp/Drawing/HorizontalRuleFormat.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldEnd.h>
#include <Aspose.Words.Cpp/Fields/FormField.h>
#include <Aspose.Words.Cpp/Fields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/HeightRule.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Replacing/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Replacing/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Replacing/ReplaceAction.h>
#include <Aspose.Words.Cpp/Replacing/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Tables/AutoFitBehavior.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Tables/CellVerticalAlignment.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/TextOrientation.h>
#include <Aspose.Words.Cpp/Underline.h>
#include <drawing/color.h>
#include <system/array.h>
#include <system/exceptions.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <testing/test_predicates.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document {

class AddContentUsingDocumentBuilder : public DocsExamplesBase
{
public:
    void CreateNewDocument()
    {
        //ExStart:CreateNewDocument
        auto doc = MakeObject<Document>();

        // Use a document builder to add content to the document.
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello World!");

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.CreateNewDocument.docx");
        //ExEnd:CreateNewDocument
    }

    void DocumentBuilderInsertBookmark()
    {
        //ExStart:DocumentBuilderInsertBookmark
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartBookmark(u"FineBookmark");
        builder->Writeln(u"This is just a fine bookmark.");
        builder->EndBookmark(u"FineBookmark");

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.DocumentBuilderInsertBookmark.docx");
        //ExEnd:DocumentBuilderInsertBookmark
    }

    void BuildTable()
    {
        //ExStart:BuildTable
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        table->AutoFit(AutoFitBehavior::FixedColumnWidths);

        builder->get_CellFormat()->set_VerticalAlignment(CellVerticalAlignment::Center);
        builder->Write(u"This is row 1 cell 1");

        builder->InsertCell();
        builder->Write(u"This is row 1 cell 2");

        builder->EndRow();

        builder->InsertCell();

        builder->get_RowFormat()->set_Height(100);
        builder->get_RowFormat()->set_HeightRule(HeightRule::Exactly);
        builder->get_CellFormat()->set_Orientation(TextOrientation::Upward);
        builder->Writeln(u"This is row 2 cell 1");

        builder->InsertCell();
        builder->get_CellFormat()->set_Orientation(TextOrientation::Downward);
        builder->Writeln(u"This is row 2 cell 2");

        builder->EndRow();
        builder->EndTable();

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.BuildTable.docx");
        //ExEnd:BuildTable
    }

    void InsertHorizontalRule()
    {
        //ExStart:InsertHorizontalRule
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Insert a horizontal rule shape into the document.");
        builder->InsertHorizontalRule();

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.InsertHorizontalRule.docx");
        //ExEnd:InsertHorizontalRule
    }

    void HorizontalRuleFormat_()
    {
        //ExStart:HorizontalRuleFormat
        auto builder = MakeObject<DocumentBuilder>();

        SharedPtr<Shape> shape = builder->InsertHorizontalRule();

        SharedPtr<HorizontalRuleFormat> horizontalRuleFormat = shape->get_HorizontalRuleFormat();
        horizontalRuleFormat->set_Alignment(HorizontalRuleAlignment::Center);
        horizontalRuleFormat->set_WidthPercent(70);
        horizontalRuleFormat->set_Height(3);
        horizontalRuleFormat->set_Color(System::Drawing::Color::get_Blue());
        horizontalRuleFormat->set_NoShade(true);

        builder->get_Document()->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.HorizontalRuleFormat.docx");
        //ExEnd:HorizontalRuleFormat
    }

    void InsertBreak()
    {
        //ExStart:InsertBreak
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"This is page 1.");
        builder->InsertBreak(BreakType::PageBreak);

        builder->Writeln(u"This is page 2.");
        builder->InsertBreak(BreakType::PageBreak);

        builder->Writeln(u"This is page 3.");

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.InsertBreak.docx");
        //ExEnd:InsertBreak
    }

    void InsertTextInputFormField()
    {
        //ExStart:InsertTextInputFormField
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u"Hello", 0);

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.InsertTextInputFormField.docx");
        //ExEnd:InsertTextInputFormField
    }

    void InsertCheckBoxFormField()
    {
        //ExStart:InsertCheckBoxFormField
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertCheckBox(u"CheckBox", true, true, 0);

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.InsertCheckBoxFormField.docx");
        //ExEnd:InsertCheckBoxFormField
    }

    void InsertComboBoxFormField()
    {
        //ExStart:InsertComboBoxFormField
        ArrayPtr<String> items = MakeArray<String>({u"One", u"Two", u"Three"});

        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertComboBox(u"DropDown", items, 0);

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.InsertComboBoxFormField.docx");
        //ExEnd:InsertComboBoxFormField
    }

    void InsertHtml()
    {
        //ExStart:InsertHtml
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertHtml(String(u"<P align='right'>Paragraph right</P>") + u"<b>Implicit paragraph left</b>" + u"<div align='center'>Div center</div>" +
                            u"<h1 align='left'>Heading 1 left.</h1>");

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.InsertHtml.docx");
        //ExEnd:InsertHtml
    }

    void InsertHyperlink()
    {
        //ExStart:InsertHyperlink
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Please make sure to visit ");
        builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
        builder->get_Font()->set_Underline(Underline::Single);

        builder->InsertHyperlink(u"Aspose Website", u"http://www.aspose.com", false);

        builder->get_Font()->ClearFormatting();
        builder->Write(u" for more information.");

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.InsertHyperlink.docx");
        //ExEnd:InsertHyperlink
    }

    void InsertTableOfContents()
    {
        //ExStart:InsertTableOfContents
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertTableOfContents(u"\\o \"1-3\" \\h \\z \\u");

        // Start the actual document content on the second page.
        builder->InsertBreak(BreakType::PageBreak);

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);

        builder->Writeln(u"Heading 1");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

        builder->Writeln(u"Heading 1.1");
        builder->Writeln(u"Heading 1.2");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);

        builder->Writeln(u"Heading 2");
        builder->Writeln(u"Heading 3");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

        builder->Writeln(u"Heading 3.1");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading3);

        builder->Writeln(u"Heading 3.1.1");
        builder->Writeln(u"Heading 3.1.2");
        builder->Writeln(u"Heading 3.1.3");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

        builder->Writeln(u"Heading 3.2");
        builder->Writeln(u"Heading 3.3");

        //ExStart:UpdateFields
        // The newly inserted table of contents will be initially empty.
        // It needs to be populated by updating the fields in the document.
        doc->UpdateFields();
        //ExEnd:UpdateFields

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.InsertTableOfContents.docx");
        //ExEnd:InsertTableOfContents
    }

    void InsertInlineImage()
    {
        //ExStart:InsertInlineImage
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertImage(ImagesDir + u"Transparent background logo.png");

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.InsertInlineImage.docx");
        //ExEnd:InsertInlineImage
    }

    void InsertFloatingImage()
    {
        //ExStart:InsertFloatingImage
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertImage(ImagesDir + u"Transparent background logo.png", RelativeHorizontalPosition::Margin, 100.0, RelativeVerticalPosition::Margin, 100.0,
                             200.0, 100.0, WrapType::Square);

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.InsertFloatingImage.docx");
        //ExEnd:InsertFloatingImage
    }

    void InsertParagraph()
    {
        //ExStart:InsertParagraph
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Font> font = builder->get_Font();
        font->set_Size(16);
        font->set_Bold(true);
        font->set_Color(System::Drawing::Color::get_Blue());
        font->set_Name(u"Arial");
        font->set_Underline(Underline::Dash);

        SharedPtr<ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
        paragraphFormat->set_FirstLineIndent(8);
        paragraphFormat->set_Alignment(ParagraphAlignment::Justify);
        paragraphFormat->set_KeepTogether(true);

        builder->Writeln(u"A whole paragraph.");

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.InsertParagraph.docx");
        //ExEnd:InsertParagraph
    }

    void InsertTCField()
    {
        //ExStart:InsertTCField
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u"TC \"Entry Text\" \\f t");

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.InsertTCField.docx");
        //ExEnd:InsertTCField
    }

    void InsertTCFieldsAtText()
    {
        //ExStart:InsertTCFieldsAtText
        auto doc = MakeObject<Document>();

        auto options = MakeObject<FindReplaceOptions>();
        options->get_ApplyFont()->set_HighlightColor(System::Drawing::Color::get_DarkOrange());
        options->set_ReplacingCallback(MakeObject<AddContentUsingDocumentBuilder::InsertTCFieldHandler>(u"Chapter 1", u"\\l 1"));

        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"The Beginning"), u"", options);
        //ExEnd:InsertTCFieldsAtText
    }

    //ExStart:InsertTCFieldHandler
    class InsertTCFieldHandler final : public IReplacingCallback
    {
    public:
        /// <summary>
        /// The display text and switches to use for each TC field. Display name can be an empty string or null.
        /// </summary>
        InsertTCFieldHandler(String text, String switches)
        {
            mFieldText = text;
            mFieldSwitches = switches;
        }

    private:
        String mFieldText;
        String mFieldSwitches;

        ReplaceAction Replacing(SharedPtr<ReplacingArgs> args) override
        {
            auto builder = MakeObject<DocumentBuilder>(System::DynamicCast<Document>(args->get_MatchNode()->get_Document()));
            builder->MoveTo(args->get_MatchNode());

            // If the user-specified text to be used in the field as display text, then use that,
            // otherwise use the match string as the display text.
            String insertText = !String::IsNullOrEmpty(mFieldText) ? mFieldText : args->get_Match()->get_Value();

            builder->InsertField(String::Format(u"TC \"{0}\" {1}", insertText, mFieldSwitches));

            return ReplaceAction::Skip;
        }
    };
    //ExEnd:InsertTCFieldHandler

    void CursorPosition()
    {
        //ExStart:CursorPosition
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Node> curNode = builder->get_CurrentNode();
        SharedPtr<Paragraph> curParagraph = builder->get_CurrentParagraph();
        //ExEnd:CursorPosition

        std::cout << (String(u"\nCursor move to paragraph: ") + curParagraph->GetText()) << std::endl;
    }

    void MoveToNode()
    {
        //ExStart:MoveToNode
        //ExStart:MoveToBookmark
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Start a bookmark and add content to it using a DocumentBuilder.
        builder->StartBookmark(u"MyBookmark");
        builder->Writeln(u"Bookmark contents.");
        builder->EndBookmark(u"MyBookmark");

        // The node that the DocumentBuilder is currently at is past the boundaries of the bookmark.
        ASPOSE_ASSERT_EQ(doc->get_Range()->get_Bookmarks()->idx_get(0)->get_BookmarkEnd(), builder->get_CurrentParagraph()->get_FirstChild());

        // If we wish to revise the content of our bookmark with the DocumentBuilder, we can move back to it like this.
        builder->MoveToBookmark(u"MyBookmark");

        // Now we're located between the bookmark's BookmarkStart and BookmarkEnd nodes, so any text the builder adds will be within it.
        ASPOSE_ASSERT_EQ(doc->get_Range()->get_Bookmarks()->idx_get(0)->get_BookmarkStart(), builder->get_CurrentParagraph()->get_FirstChild());

        // We can move the builder to an individual node,
        // which in this case will be the first node of the first paragraph, like this.
        builder->MoveTo(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetChildNodes(NodeType::Any, false)->idx_get(0));
        //ExEnd:MoveToBookmark

        ASSERT_EQ(NodeType::BookmarkStart, builder->get_CurrentNode()->get_NodeType());
        ASSERT_TRUE(builder->get_IsAtStartOfParagraph());

        // A shorter way of moving the very start/end of a document is with these methods.
        builder->MoveToDocumentEnd();
        ASSERT_TRUE(builder->get_IsAtEndOfParagraph());
        builder->MoveToDocumentStart();
        ASSERT_TRUE(builder->get_IsAtStartOfParagraph());
        //ExEnd:MoveToNode
    }

    void MoveToDocumentStartEnd()
    {
        //ExStart:MoveToDocumentStartEnd
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Move the cursor position to the beginning of your document.
        builder->MoveToDocumentStart();
        std::cout << "\nThis is the beginning of the document." << std::endl;

        // Move the cursor position to the end of your document.
        builder->MoveToDocumentEnd();
        std::cout << "\nThis is the end of the document." << std::endl;
        //ExEnd:MoveToDocumentStartEnd
    }

    void MoveToSection()
    {
        //ExStart:MoveToSection
        auto doc = MakeObject<Document>();
        doc->AppendChild(MakeObject<Section>(doc));

        // Move a DocumentBuilder to the second section and add text.
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveToSection(1);
        builder->Writeln(u"Text added to the 2nd section.");

        // Create document with paragraphs.
        doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
        ASSERT_EQ(23, paragraphs->get_Count());

        // When we create a DocumentBuilder for a document, its cursor is at the very beginning of the document by default,
        // and any content added by the DocumentBuilder will just be prepended to the document.
        builder = MakeObject<DocumentBuilder>(doc);
        ASSERT_EQ(0, paragraphs->IndexOf(builder->get_CurrentParagraph()));

        // You can move the cursor to any position in a paragraph.
        builder->MoveToParagraph(2, 10);
        ASSERT_EQ(2, paragraphs->IndexOf(builder->get_CurrentParagraph()));
        builder->Writeln(u"This is a new third paragraph. ");
        ASSERT_EQ(3, paragraphs->IndexOf(builder->get_CurrentParagraph()));
        //ExEnd:MoveToSection
    }

    void MoveToHeadersFooters()
    {
        //ExStart:MoveToHeadersFooters
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Specify that we want headers and footers different for first, even and odd pages.
        builder->get_PageSetup()->set_DifferentFirstPageHeaderFooter(true);
        builder->get_PageSetup()->set_OddAndEvenPagesHeaderFooter(true);

        // Create the headers.
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderFirst);
        builder->Write(u"Header for the first page");
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderEven);
        builder->Write(u"Header for even pages");
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Write(u"Header for all other pages");

        // Create two pages in the document.
        builder->MoveToSection(0);
        builder->Writeln(u"Page1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page2");

        doc->Save(ArtifactsDir + u"AddContentUsingDocumentBuilder.MoveToHeadersFooters.docx");
        //ExEnd:MoveToHeadersFooters
    }

    void MoveToParagraph()
    {
        //ExStart:MoveToParagraph
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->MoveToParagraph(2, 0);
        builder->Writeln(u"This is the 3rd paragraph.");
        //ExEnd:MoveToParagraph
    }

    void MoveToTableCell()
    {
        //ExStart:MoveToTableCell
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Move the builder to row 3, cell 4 of the first table.
        builder->MoveToCell(0, 2, 3, 0);
        builder->Write(u"\nCell contents added by DocumentBuilder");
        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        ASPOSE_ASSERT_EQ(table->get_Rows()->idx_get(2)->get_Cells()->idx_get(3), builder->get_CurrentNode()->get_ParentNode()->get_ParentNode());
        ASSERT_EQ(u"Cell contents added by DocumentBuilderCell 3 contents\a", table->get_Rows()->idx_get(2)->get_Cells()->idx_get(3)->GetText().Trim());
        //ExEnd:MoveToTableCell
    }

    void MoveToBookmarkEnd()
    {
        //ExStart:MoveToBookmarkEnd
        auto doc = MakeObject<Document>(MyDir + u"Bookmarks.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->MoveToBookmark(u"MyBookmark1", false, true);
        builder->Writeln(u"This is a bookmark.");
        //ExEnd:MoveToBookmarkEnd
    }

    void MoveToMergeField()
    {
        //ExStart:MoveToMergeField
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a field using the DocumentBuilder and add a run of text after it.
        SharedPtr<Field> field = builder->InsertField(u"MERGEFIELD field");
        builder->Write(u" Text after the field.");

        // The builder's cursor is currently at end of the document.
        ASSERT_TRUE(builder->get_CurrentNode() == nullptr);
        // We can move the builder to a field like this, placing the cursor at immediately after the field.
        builder->MoveToField(field, true);

        // Note that the cursor is at a place past the FieldEnd node of the field, meaning that we are not actually inside the field.
        // If we wish to move the DocumentBuilder to inside a field,
        // we will need to move it to a field's FieldStart or FieldSeparator node using the DocumentBuilder.MoveTo() method.
        ASPOSE_ASSERT_EQ(field->get_End(), builder->get_CurrentNode()->get_PreviousSibling());
        builder->Write(u" Text immediately after the field.");
        //ExEnd:MoveToMergeField
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document
