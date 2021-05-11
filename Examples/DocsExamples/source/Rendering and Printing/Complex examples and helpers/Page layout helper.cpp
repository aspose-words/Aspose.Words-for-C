#include "Rendering and Printing/Complex examples and helpers/Page layout helper.h"

#include <iostream>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Range.h>
#include <system/enumerator_adapter.h>
#include <system/environment.h>
#include <system/exceptions.h>
#include <system/scope_guard.h>
#include <system/text/string_builder.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;
namespace DocsExamples { namespace Complex_examples_and_helpers {

namespace gtest_test {

class DocumentLayoutHelper : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Complex_examples_and_helpers::DocumentLayoutHelper> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Complex_examples_and_helpers::DocumentLayoutHelper>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Complex_examples_and_helpers::DocumentLayoutHelper> DocumentLayoutHelper::s_instance;

} // namespace gtest_test

void DocumentLayoutHelper::WrapperToAccessLayoutEntities()
{
    // This sample introduces the RenderedDocument class and other related classes which provide an API wrapper for
    // the LayoutEnumerator. This allows you to access the layout entities of a document using a DOM style API.
    auto doc = System::MakeObject<Document>(MyDir + u"Document layout.docx");

    auto layoutDoc = System::MakeObject<RenderedDocument>(doc);

    // Get access to the line of the first page and print to the console.
    System::SharedPtr<RenderedLine> line = layoutDoc->get_Pages()->idx_get(0)->get_Columns()->idx_get(0)->get_Lines()->idx_get(2);
    std::cout << (System::String(u"Line: ") + line->get_Text()) << std::endl;

    // With a rendered line, the original paragraph in the document object model can be returned.
    System::SharedPtr<Paragraph> para = line->get_Paragraph();
    std::cout << (System::String(u"Paragraph text: ") + para->get_Range()->get_Text()) << std::endl;

    // Retrieve all the text that appears on the first page in plain text format (including headers and footers).
    System::String pageText = layoutDoc->get_Pages()->idx_get(0)->get_Text();
    std::cout << std::endl;

    // Loop through each page in the document and print how many lines appear on each page.
    for (auto page : System::IterateOver(layoutDoc->get_Pages()))
    {
        System::SharedPtr<LayoutCollection<System::SharedPtr<LayoutEntity>>> lines = page->GetChildEntities(LayoutEntityType::Line, true);
        std::cout << "Page " << page->get_PageIndex() << " has " << lines->get_Count() << " lines." << std::endl;
    }

    // This method provides a reverse lookup of layout entities for any given node
    // (except runs and nodes in the header and footer).
    std::cout << std::endl;
    std::cout << "The lines of the second paragraph:" << std::endl;
    for (auto paragraphLine :
         System::IterateOver<RenderedLine>(layoutDoc->GetLayoutEntitiesOfNode(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1))))
    {
        std::cout << "\"" << paragraphLine->get_Text().Trim() << "\"" << std::endl;
        std::cout << System::ObjectExt::ToString(paragraphLine->get_Rectangle()) << std::endl;
        std::cout << std::endl;
    }
}

namespace gtest_test {

TEST_F(DocumentLayoutHelper, WrapperToAccessLayoutEntities)
{
    s_instance->WrapperToAccessLayoutEntities();
}

} // namespace gtest_test

int32_t LayoutEntity::get_PageIndex()
{
    return mPageIndex;
}

System::Drawing::RectangleF LayoutEntity::get_Rectangle()
{
    return mRectangle;
}

LayoutEntityType LayoutEntity::get_Type()
{
    return mType;
}

System::String LayoutEntity::get_Text()
{
    auto builder = System::MakeObject<System::Text::StringBuilder>();
    for (auto entity : mChildEntities)
    {
        builder->Append(entity->get_Text());
    }

    return builder->ToString();
}

System::SharedPtr<LayoutEntity> LayoutEntity::get_Parent()
{
    return mParent;
}

System::SharedPtr<Node> LayoutEntity::get_ParentNode()
{
    return mParentNode;
}

System::SharedPtr<System::Object> LayoutEntity::get_LayoutObject()
{
    return pr_LayoutObject;
}

void LayoutEntity::set_LayoutObject(System::SharedPtr<System::Object> value)
{
    pr_LayoutObject = value;
}

void LayoutEntity::SetParentNode(System::SharedPtr<Node> value)
{
    mParentNode = value;
}

System::SharedPtr<LayoutEntity> LayoutEntity::AddChildEntity(System::SharedPtr<LayoutEnumerator> it)
{
    System::SharedPtr<LayoutEntity> child = CreateLayoutEntityFromType(it);
    mChildEntities->Add(child);

    return child;
}

System::SharedPtr<LayoutEntity> LayoutEntity::CreateLayoutEntityFromType(System::SharedPtr<LayoutEnumerator> it)
{
    System::SharedPtr<LayoutEntity> childEntity;
    switch (it->get_Type())
    {
    case LayoutEntityType::Cell:
        childEntity = System::MakeObject<RenderedCell>();
        break;

    case LayoutEntityType::Column:
        childEntity = System::MakeObject<RenderedColumn>();
        break;

    case LayoutEntityType::Comment:
        childEntity = System::MakeObject<RenderedComment>();
        break;

    case LayoutEntityType::Endnote:
        childEntity = System::MakeObject<RenderedEndnote>();
        break;

    case LayoutEntityType::Footnote:
        childEntity = System::MakeObject<RenderedFootnote>();
        break;

    case LayoutEntityType::HeaderFooter:
        childEntity = System::MakeObject<RenderedHeaderFooter>();
        break;

    case LayoutEntityType::Line:
        childEntity = System::MakeObject<RenderedLine>();
        break;

    case LayoutEntityType::NoteSeparator:
        childEntity = System::MakeObject<RenderedNoteSeparator>();
        break;

    case LayoutEntityType::Page:
        childEntity = System::MakeObject<RenderedPage>();
        break;

    case LayoutEntityType::Row:
        childEntity = System::MakeObject<RenderedRow>();
        break;

    case LayoutEntityType::Span:
        childEntity = RenderedSpan::MakeObject(it->get_Text());
        break;

    case LayoutEntityType::TextBox:
        childEntity = System::MakeObject<RenderedTextBox>();
        break;

    default:
        throw System::InvalidOperationException(u"Unknown layout type");
    }

    childEntity->mKind = it->get_Kind();
    childEntity->mPageIndex = it->get_PageIndex();
    childEntity->mRectangle = it->get_Rectangle();
    childEntity->mType = it->get_Type();
    childEntity->set_LayoutObject(it->get_Current());
    childEntity->mParent = System::MakeSharedPtr(this);

    return childEntity;
}

System::SharedPtr<LayoutCollection<System::SharedPtr<LayoutEntity>>> LayoutEntity::GetChildEntities(LayoutEntityType type, bool isDeep)
{
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<LayoutEntity>>> childList =
        System::MakeObject<System::Collections::Generic::List<System::SharedPtr<LayoutEntity>>>();

    for (auto entity : mChildEntities)
    {
        if ((entity->get_Type() & type) == entity->get_Type())
        {
            childList->Add(entity);
        }

        if (isDeep)
        {
            childList->AddRange(entity->GetChildEntities(type, true));
        }
    }

    return LayoutCollection<System::SharedPtr<LayoutEntity>>::MakeObject(childList);
}

LayoutEntity::LayoutEntity()
    : mPageIndex(0), mType(((Aspose::Words::Layout::LayoutEntityType)0)),
      mChildEntities(System::MakeObject<System::Collections::Generic::List<System::SharedPtr<LayoutEntity>>>())
{
}

LayoutEntity::~LayoutEntity()
{
}

System::SharedPtr<LayoutCollection<System::SharedPtr<RenderedPage>>> RenderedDocument::get_Pages()
{
    return GetChildNodes<System::SharedPtr<RenderedPage>>();
}

RenderedDocument::RenderedDocument(System::SharedPtr<Document> doc)
    : mLayoutToNodeLookup(System::MakeObject<System::Collections::Generic::Dictionary<System::SharedPtr<System::Object>, System::SharedPtr<Node>>>())
{
    // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference

    mLayoutCollector = System::MakeObject<LayoutCollector>(doc);
    mEnumerator = System::MakeObject<LayoutEnumerator>(doc);

    ProcessLayoutElements(System::MakeSharedPtr(this));
    LinkLayoutMarkersToNodes(doc);
    CollectLinesAndAddToMarkers();
}

System::SharedPtr<LayoutCollection<System::SharedPtr<LayoutEntity>>> RenderedDocument::GetLayoutEntitiesOfNode(System::SharedPtr<Node> node)
{
    if (!System::ObjectExt::Equals(mLayoutCollector->get_Document(), node->get_Document()))
    {
        throw System::ArgumentException(u"Node does not belong to the same document which was rendered.");
    }

    if (node->get_NodeType() == NodeType::Document)
    {
        return LayoutCollection<System::SharedPtr<LayoutEntity>>::MakeObject(mChildEntities);
    }

    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<LayoutEntity>>> entities =
        System::MakeObject<System::Collections::Generic::List<System::SharedPtr<LayoutEntity>>>();

    // Retrieve all entities from the layout document (inversion of LayoutEntityType.None).
    for (auto entity : System::IterateOver(GetChildEntities(~LayoutEntityType::None, true)))
    {
        if (entity->get_ParentNode() == node)
        {
            entities->Add(entity);
        }

        // There is no table entity in rendered output, so manually check if rows belong to a table node.
        if (entity->get_Type() == LayoutEntityType::Row)
        {
            auto row = System::StaticCast<RenderedRow>(entity);
            if (row->get_Table() == node)
            {
                entities->Add(entity);
            }
        }
    }

    return LayoutCollection<System::SharedPtr<LayoutEntity>>::MakeObject(entities);
}

void RenderedDocument::ProcessLayoutElements(System::SharedPtr<LayoutEntity> current)
{
    do
    {
        System::SharedPtr<LayoutEntity> child = current->AddChildEntity(mEnumerator);

        if (mEnumerator->MoveFirstChild())
        {
            current = child;

            ProcessLayoutElements(current);
            mEnumerator->MoveParent();

            current = current->get_Parent();
        }
    } while (mEnumerator->MoveNext());
}

void RenderedDocument::CollectLinesAndAddToMarkers()
{
    CollectLinesOfMarkersCore(LayoutEntityType::Column);
    CollectLinesOfMarkersCore(LayoutEntityType::Comment);
}

void RenderedDocument::CollectLinesOfMarkersCore(LayoutEntityType type)
{
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<RenderedLine>>> collectedLines =
        System::MakeObject<System::Collections::Generic::List<System::SharedPtr<RenderedLine>>>();

    for (auto page : System::IterateOver(get_Pages()))
    {
        for (auto story : System::IterateOver(page->GetChildEntities(type, false)))
        {
            for (auto line : System::IterateOver<RenderedLine>(story->GetChildEntities(LayoutEntityType::Line, true)))
            {
                collectedLines->Add(line);
                for (auto span : System::IterateOver(line->get_Spans()))
                {
                    if (mLayoutToNodeLookup->ContainsKey(span->get_LayoutObject()))
                    {
                        if (span->get_Kind() == u"PARAGRAPH" || span->get_Kind() == u"ROW" || span->get_Kind() == u"CELL" || span->get_Kind() == u"SECTION")
                        {
                            System::SharedPtr<Node> node = mLayoutToNodeLookup->idx_get(span->get_LayoutObject());

                            if (node->get_NodeType() == NodeType::Row)
                            {
                                node = (System::DynamicCast<Row>(node))->get_LastCell()->get_LastParagraph();
                            }

                            for (auto collectedLine : collectedLines)
                            {
                                collectedLine->SetParentNode(node);
                            }

                            collectedLines = System::MakeObject<System::Collections::Generic::List<System::SharedPtr<RenderedLine>>>();
                        }
                        else
                        {
                            span->SetParentNode(mLayoutToNodeLookup->idx_get(span->get_LayoutObject()));
                        }
                    }
                }
            }
        }
    }
}

void RenderedDocument::LinkLayoutMarkersToNodes(System::SharedPtr<Document> doc)
{
    for (auto node : System::IterateOver(doc->GetChildNodes(NodeType::Any, true)))
    {
        System::SharedPtr<System::Object> entity = mLayoutCollector->GetEntity(node);

        if (entity != nullptr)
        {
            mLayoutToNodeLookup->Add(entity, node);
        }
    }
}

RenderedDocument::~RenderedDocument()
{
}

System::SharedPtr<LayoutCollection<System::SharedPtr<RenderedLine>>> StoryLayoutEntity::get_Lines()
{
    return GetChildNodes<System::SharedPtr<RenderedLine>>();
}

System::SharedPtr<LayoutCollection<System::SharedPtr<RenderedRow>>> StoryLayoutEntity::get_Rows()
{
    return GetChildNodes<System::SharedPtr<RenderedRow>>();
}

System::String RenderedLine::get_Text()
{
    return DocsExamples::Complex_examples_and_helpers::LayoutEntity::get_Text() + System::Environment::get_NewLine();
}

System::SharedPtr<Aspose::Words::Paragraph> RenderedLine::get_Paragraph()
{
    return System::DynamicCast<Aspose::Words::Paragraph>(get_ParentNode());
}

System::SharedPtr<LayoutCollection<System::SharedPtr<RenderedSpan>>> RenderedLine::get_Spans()
{
    return GetChildNodes<System::SharedPtr<RenderedSpan>>();
}

System::String RenderedSpan::get_Kind()
{
    return mKind;
}

System::String RenderedSpan::get_Text()
{
    return pr_Text;
}

System::SharedPtr<Node> RenderedSpan::get_ParentNode()
{
    return mParentNode;
}

RenderedSpan::RenderedSpan()
{
}

RenderedSpan::RenderedSpan(System::String text)
{
    // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference

    // Assign empty text if the span text is null (this can happen with shape spans).
    pr_Text = text != nullptr ? text : System::String::Empty;
}

MEMBER_FUNCTION_MAKE_OBJECT_DEFINITION(RenderedSpan, CODEPORTING_ARGS(System::String text), CODEPORTING_ARGS(text));

System::String RenderedHeaderFooter::get_Kind()
{
    return mKind;
}

System::SharedPtr<LayoutCollection<System::SharedPtr<RenderedColumn>>> RenderedPage::get_Columns()
{
    return GetChildNodes<System::SharedPtr<RenderedColumn>>();
}

System::SharedPtr<LayoutCollection<System::SharedPtr<RenderedHeaderFooter>>> RenderedPage::get_HeaderFooters()
{
    return GetChildNodes<System::SharedPtr<RenderedHeaderFooter>>();
}

System::SharedPtr<LayoutCollection<System::SharedPtr<RenderedComment>>> RenderedPage::get_Comments()
{
    return GetChildNodes<System::SharedPtr<RenderedComment>>();
}

System::SharedPtr<Aspose::Words::Section> RenderedPage::get_Section()
{
    return System::DynamicCast<Aspose::Words::Section>(get_ParentNode());
}

System::SharedPtr<Node> RenderedPage::get_ParentNode()
{
    return get_Columns()->get_First()->GetChildEntities(LayoutEntityType::Line, true)->get_First()->get_ParentNode()->GetAncestor(NodeType::Section);
}

System::SharedPtr<LayoutCollection<System::SharedPtr<RenderedCell>>> RenderedRow::get_Cells()
{
    return GetChildNodes<System::SharedPtr<RenderedCell>>();
}

System::SharedPtr<Aspose::Words::Tables::Row> RenderedRow::get_Row()
{
    return System::DynamicCast<Aspose::Words::Tables::Row>(get_ParentNode());
}

System::SharedPtr<Aspose::Words::Tables::Table> RenderedRow::get_Table()
{
    return get_Row() != nullptr ? get_Row()->get_ParentTable() : nullptr;
}

System::SharedPtr<Node> RenderedRow::get_ParentNode()
{
    System::SharedPtr<Paragraph> para =
        get_Cells()->get_First()->get_Lines()->get_First() != nullptr ? get_Cells()->get_First()->get_Lines()->get_First()->get_Paragraph() : nullptr;
    return para != nullptr ? System::StaticCast<Node>(para->GetAncestor(NodeType::Row)) : nullptr;
}

System::SharedPtr<LayoutCollection<System::SharedPtr<RenderedFootnote>>> RenderedColumn::get_Footnotes()
{
    return GetChildNodes<System::SharedPtr<RenderedFootnote>>();
}

System::SharedPtr<LayoutCollection<System::SharedPtr<RenderedEndnote>>> RenderedColumn::get_Endnotes()
{
    return GetChildNodes<System::SharedPtr<RenderedEndnote>>();
}

System::SharedPtr<LayoutCollection<System::SharedPtr<RenderedNoteSeparator>>> RenderedColumn::get_NoteSeparators()
{
    return GetChildNodes<System::SharedPtr<RenderedNoteSeparator>>();
}

System::SharedPtr<Aspose::Words::Body> RenderedColumn::get_Body()
{
    return System::DynamicCast<Aspose::Words::Body>(get_ParentNode());
}

System::SharedPtr<Node> RenderedColumn::get_ParentNode()
{
    return GetChildEntities(LayoutEntityType::Line, true)->get_First()->get_ParentNode()->GetAncestor(NodeType::Body);
}

System::SharedPtr<Aspose::Words::Tables::Cell> RenderedCell::get_Cell()
{
    return System::DynamicCast<Aspose::Words::Tables::Cell>(get_ParentNode());
}

System::SharedPtr<Node> RenderedCell::get_ParentNode()
{
    auto firstLine = get_Lines()->get_First();
    if (firstLine != nullptr)
    {
        return firstLine->get_Paragraph() != nullptr ? System::StaticCast<Node>(firstLine->get_Paragraph()->GetAncestor(NodeType::Cell)) : nullptr;
    }
    return nullptr;
}

System::SharedPtr<Aspose::Words::Notes::Footnote> RenderedFootnote::get_Footnote()
{
    return System::DynamicCast<Aspose::Words::Notes::Footnote>(get_ParentNode());
}

System::SharedPtr<Node> RenderedFootnote::get_ParentNode()
{
    return GetChildEntities(LayoutEntityType::Line, true)->get_First()->get_ParentNode()->GetAncestor(NodeType::Footnote);
}

System::SharedPtr<Footnote> RenderedEndnote::get_Endnote()
{
    return System::DynamicCast<Footnote>(get_ParentNode());
}

System::SharedPtr<Node> RenderedEndnote::get_ParentNode()
{
    return GetChildEntities(LayoutEntityType::Line, true)->get_First()->get_ParentNode()->GetAncestor(NodeType::Footnote);
}

System::SharedPtr<Node> RenderedTextBox::get_ParentNode()
{
    System::SharedPtr<LayoutCollection<System::SharedPtr<LayoutEntity>>> lines = GetChildEntities(LayoutEntityType::Line, true);
    System::SharedPtr<Node> shape = lines->get_First()->get_ParentNode()->GetAncestor(NodeType::Shape);

    return shape != nullptr ? shape : lines->get_First()->get_ParentNode()->GetAncestor(NodeType::Shape);
}

System::SharedPtr<Aspose::Words::Comment> RenderedComment::get_Comment()
{
    return System::DynamicCast<Aspose::Words::Comment>(get_ParentNode());
}

System::SharedPtr<Node> RenderedComment::get_ParentNode()
{
    return GetChildEntities(LayoutEntityType::Line, true)->get_First()->get_ParentNode()->GetAncestor(NodeType::Comment);
}

System::SharedPtr<Aspose::Words::Notes::Footnote> RenderedNoteSeparator::get_Footnote()
{
    return System::DynamicCast<Aspose::Words::Notes::Footnote>(get_ParentNode());
}

System::SharedPtr<Node> RenderedNoteSeparator::get_ParentNode()
{
    return GetChildEntities(LayoutEntityType::Line, true)->get_First()->get_ParentNode()->GetAncestor(NodeType::Footnote);
}

}} // namespace DocsExamples::Complex_examples_and_helpers
