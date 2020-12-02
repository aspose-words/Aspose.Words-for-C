#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Comparer/CompareOptions.h>
#include <Aspose.Words.Cpp/Model/Comparer/ComparisonTargetType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionCollection.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldDate.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
namespace
{
    void NormalComparison(System::String const &inputDataDir)
    {
        // ExStart:NormalComparison
        System::SharedPtr<Document> docA = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
        System::SharedPtr<Document> docB = System::MakeObject<Document>(inputDataDir + u"TestFile - Copy.doc");
        // DocA now contains changes as revisions. 
        docA->Compare(docB, u"user", System::DateTime::get_Now());
        // ExEnd:NormalComparison
    }

    void CompareForEqual(System::String const &inputDataDir)
    {
        // ExStart:CompareForEqual
        System::SharedPtr<Document> docA = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
        System::SharedPtr<Document> docB = System::MakeObject<Document>(inputDataDir + u"TestFile - Copy.doc");
        // DocA now contains changes as revisions. 
        docA->Compare(docB, u"user", System::DateTime::get_Now());
        if (docA->get_Revisions()->get_Count() == 0)
        {
            std::cout << "Documents are equal" << std::endl;
        }
        else
        {
            std::cout << "Documents are not equal" << std::endl;
        }
        // ExEnd:CompareForEqual
    }

    void CompareDocumentWithCompareOptions(System::String const &inputDataDir)
    {
        // ExStart:CompareDocumentWithCompareOptions
        // Create the original document
        System::SharedPtr<Document> docOriginal = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(docOriginal);

        // Insert paragraph text with an endnote
        builder->Writeln(u"Hello world! This is the first paragraph.");
        builder->InsertFootnote(FootnoteType::Endnote, u"Original endnote text.");

        // Insert a table
        builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Original cell 1 text");
        builder->InsertCell();
        builder->Write(u"Original cell 2 text");
        builder->EndTable();

        // Insert a textbox
        System::SharedPtr<Shape> textBox = builder->InsertShape(ShapeType::TextBox, 150, 20);
        builder->MoveTo(textBox->get_FirstParagraph());
        builder->Write(u"Original textbox contents");

        // Insert a DATE field
        builder->MoveTo(docOriginal->get_FirstSection()->get_Body()->AppendParagraph(u""));
        builder->InsertField(u" DATE ");

        // Insert a comment
        auto newComment = System::MakeObject<Comment>(docOriginal, u"John Doe", u"J.D.", System::DateTime::get_Now());
        newComment->SetText(u"Original comment.");
        builder->get_CurrentParagraph()->AppendChild(newComment);

        // Insert a header
        builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
        builder->Writeln(u"Original header contents.");

        // Create a clone of our document, which we will edit and later compare to the original
        auto docEdited = System::DynamicCast<Aspose::Words::Document>(System::StaticCast<Node>(docOriginal)->Clone(true));
        System::SharedPtr<Paragraph> firstParagraph = docEdited->get_FirstSection()->get_Body()->get_FirstParagraph();

        // Change the formatting of the first paragraph, change casing of original characters and add text
        firstParagraph->get_Runs()->idx_get(0)->set_Text(u"hello world! this is the first paragraph, after editing.");
        firstParagraph->get_ParagraphFormat()->set_Style(docEdited->get_Styles()->idx_get(Aspose::Words::StyleIdentifier::Heading1));

        // Edit the footnote
        auto footnote = System::DynamicCast<Aspose::Words::Footnote>(docEdited->GetChild(Aspose::Words::NodeType::Footnote, 0, true));
        footnote->get_FirstParagraph()->get_Runs()->idx_get(1)->set_Text(u"Edited endnote text.");

        // Edit the table
        auto table = System::DynamicCast<Aspose::Words::Tables::Table>(docEdited->GetChild(Aspose::Words::NodeType::Table, 0, true));
        table->get_FirstRow()->get_Cells()->idx_get(1)->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Edited Cell 2 contents");

        // Edit the textbox
        textBox = System::DynamicCast<Aspose::Words::Drawing::Shape>(docEdited->GetChild(Aspose::Words::NodeType::Shape, 0, true));
        textBox->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Edited textbox contents");

        // Edit the DATE field
        auto fieldDate = System::DynamicCast<Aspose::Words::Fields::FieldDate>(docEdited->get_Range()->get_Fields()->idx_get(0));
        fieldDate->set_UseLunarCalendar(true);

        // Edit the comment
        auto comment = System::DynamicCast<Aspose::Words::Comment>(docEdited->GetChild(Aspose::Words::NodeType::Comment, 0, true));
        comment->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Edited comment.");

        // Edit the header
        docEdited->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Edited header contents.");

        // When we compare documents, the differences of the latter document from the former show up as revisions to the former
        // Each edit that we've made above will have its own revision, after we run the Compare method
        // We can compare with a CompareOptions object, which can suppress changes done to certain types of objects within the original document
        // from registering as revisions after the comparison by setting some of these members to "true"
        auto compareOptions = System::MakeObject<Aspose::Words::CompareOptions>();
        compareOptions->set_IgnoreFormatting(false);
        compareOptions->set_IgnoreCaseChanges(false);
        compareOptions->set_IgnoreComments(false);
        compareOptions->set_IgnoreTables(false);
        compareOptions->set_IgnoreFields(false);
        compareOptions->set_IgnoreFootnotes(false);
        compareOptions->set_IgnoreTextboxes(false);
        compareOptions->set_IgnoreHeadersAndFooters(false);
        compareOptions->set_Target(Aspose::Words::ComparisonTargetType::New);

        docOriginal->Compare(docEdited, u"John Doe", System::DateTime::get_Now(), compareOptions);
        docOriginal->Save(inputDataDir + u"Document.CompareOptions.docx");
        // ExEnd:CompareDocumentWithCompareOptions
    }

    void CompareDocumentWithComparisonTarget(System::String const &inputDataDir)
    {
        // ExStart:CompareDocumentWithComparisonTarget
        System::SharedPtr<Document> docA = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
        System::SharedPtr<Document> docB = System::MakeObject<Document>(inputDataDir + u"TestFile - Copy.doc");

        System::SharedPtr<CompareOptions> options = System::MakeObject<CompareOptions>();
        options->set_IgnoreFormatting(true);
        // Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box. 
        options->set_Target(ComparisonTargetType::New);

        docA->Compare(docB, u"user", System::DateTime::get_Now(), options);
        // ExEnd:CompareDocumentWithComparisonTarget

        std::cout << "Documents have compared successfully." << std::endl;
    }

    void SpecifyComparisonGranularity(System::String const &inputDataDir)
    {
        // ExStart:SpecifyComparisonGranularity
        System::SharedPtr<DocumentBuilder> builderA = System::MakeObject<DocumentBuilder>(System::MakeObject<Document>());
        System::SharedPtr<DocumentBuilder> builderB = System::MakeObject<DocumentBuilder>(System::MakeObject<Document>());

        builderA->Writeln(u"This is A simple word");
        builderB->Writeln(u"This is B simple words");

        System::SharedPtr<CompareOptions> co = System::MakeObject<CompareOptions>();
        co->set_Granularity(Aspose::Words::Granularity::CharLevel);

        builderA->get_Document()->Compare(builderB->get_Document(), u"author", System::DateTime::get_Now(), co);
        // ExEnd:SpecifyComparisonGranularity
    }

    void CompareWhenDocumentHasRevisions(System::String const& inputDataDir)
    {
        // ExStart:CompareWhenDocumentHasRevisions
        // The source document doc1
        System::SharedPtr<Document> doc1 = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc1);
        builder->Writeln(u"This is the original document.");

        // The target document doc2
        System::SharedPtr<Document> doc2 = System::MakeObject<Document>();
        builder = System::MakeObject<DocumentBuilder>(doc2);
        builder->Writeln(u"This is the edited document.");

        // If either document has a revision, an exception will be thrown
        if (doc1->get_Revisions()->get_Count() == 0 && doc2->get_Revisions()->get_Count() == 0)
            doc1->Compare(doc2, u"authorName", System::DateTime::get_Now());

        // If doc1 and doc2 are different, doc1 now has some revisions after the comparison, which can now be viewed and processed
        if (doc1->get_Revisions()->get_Count() == 2)
            std::cout << "Documents are equal." << std::endl << std::endl;

        for (System::SharedPtr<Revision> r : System::IterateOver(doc1->get_Revisions()))
        {
            std::cout << "Revision type: " << System::ObjectExt::ToString(r->get_RevisionType()).ToUtf8String() 
                << ", on a node of type " << System::ObjectExt::ToString(r->get_ParentNode()->get_NodeType()).ToUtf8String() << std::endl;
            std::cout << "Changed text: " << r->get_ParentNode()->GetText() << std::endl;
        }

        // All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2
        doc1->get_Revisions()->AcceptAll();

        // doc1, when saved, now resembles doc2
        doc1->Save(inputDataDir + u"Document.Compare.docx");
        doc1 = System::MakeObject<Document>(inputDataDir + u"Document.Compare.docx");

        if (doc1->get_Revisions()->get_Count() == 0)
            std::cout << "Documents are equal" << std::endl;

        if (doc2->GetText().Trim() == doc1->GetText().Trim())
            std::cout << "Documents are equal" << std::endl;

        // ExEnd:CompareWhenDocumentHasRevisions
    }
}

void CompareDocument()
{
    std::cout << "CompareDocument example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    NormalComparison(inputDataDir);
    CompareForEqual(inputDataDir);
    CompareDocumentWithCompareOptions(inputDataDir);
    CompareDocumentWithComparisonTarget(inputDataDir);
    SpecifyComparisonGranularity(inputDataDir);
    CompareWhenDocumentHasRevisions(inputDataDir);
    std::cout << "CompareDocument example finished." << std::endl << std::endl;
}