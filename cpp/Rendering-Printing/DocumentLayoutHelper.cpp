#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Layout/Public/LayoutCollector.h>
#include <Layout/Public/LayoutEntityType.h>
#include <Layout/Public/LayoutEnumerator.h>
#include <Model/Document/Document.h>
#include <Model/Footnotes/Footnote.h>
#include <Model/Nodes/Node.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/Section.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/Table.h>
#include <Model/Text/Comment.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/ParagraphCollection.h>
#include <Model/Text/Range.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Tables;

namespace
{
    template <typename T> T GetFirst(System::SharedPtr<System::Collections::Generic::IList<T>> source)
    {
        return source->get_Count() > 0 ? source->idx_get(0) : System::Default<T>();
    }

    class LayoutEntity;
    class RenderedLine;
    class RenderedSpan;
    class RenderedRow;
    class RenderedColumn;
    class RenderedCell;
    class RenderedFootnote;
    class RenderedEndnote;
    class RenderedNoteSeparator;
    class RenderedHeaderFooter;
    class RenderedComment;
    class RenderedPage;

    typedef System::Collections::Generic::IList<System::SharedPtr<LayoutEntity>> TLayoutEntityIList;
    typedef System::Collections::Generic::List<System::SharedPtr<LayoutEntity>> TLayoutEntityList;
    typedef System::Collections::Generic::IList<System::SharedPtr<RenderedLine>> TRenderedLineIList;
    typedef System::Collections::Generic::List<System::SharedPtr<RenderedLine>> TRenderedLineList;
    typedef System::Collections::Generic::IList<System::SharedPtr<RenderedRow>> TRenderedRowIList;
    typedef System::Collections::Generic::IList<System::SharedPtr<RenderedSpan>> TRenderedSpanIList;
    typedef System::Collections::Generic::IList<System::SharedPtr<RenderedCell>> TRenderedCellIList;
    typedef System::Collections::Generic::IList<System::SharedPtr<RenderedFootnote>> TRenderedFootnoteIList;
    typedef System::Collections::Generic::IList<System::SharedPtr<RenderedEndnote>> TRenderedEndnoteIList;
    typedef System::Collections::Generic::IList<System::SharedPtr<RenderedNoteSeparator>> TRenderedNoteSeparatorIList;
    typedef System::Collections::Generic::IList<System::SharedPtr<RenderedColumn>> TRenderedColumnIList;
    typedef System::Collections::Generic::IList<System::SharedPtr<RenderedHeaderFooter>> TRenderedHeaderFooterIList;
    typedef System::Collections::Generic::IList<System::SharedPtr<RenderedComment>> TRenderedCommentIList;
    typedef System::Collections::Generic::IList<System::SharedPtr<RenderedPage>> TRenderedPageIList;
    typedef System::Collections::Generic::IDictionary<System::SharedPtr<System::Object>, System::SharedPtr<Node>> LayoutNodeLookupIDict;
    typedef System::Collections::Generic::Dictionary<System::SharedPtr<System::Object>, System::SharedPtr<Node>> LayoutNodeLookupDict;

    System::SharedPtr<LayoutEntity> CreateLayoutEntityFromType(System::SharedPtr<LayoutEnumerator> it);

    /// <summary>
    /// Provides the base class for rendered elements of a document.
    /// </summary>
    class LayoutEntity : public System::Object
    {
        typedef LayoutEntity ThisType;
        typedef System::Object BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        friend class RenderedDocument;

    public:
        int32_t GetPageIndex();
        System::Drawing::RectangleF GetRectangle();
        LayoutEntityType GetLayoutEntityType();
        virtual System::String GetText();
        System::SharedPtr<LayoutEntity> GetParent();
        virtual System::SharedPtr<Node> GetParentNode();
        System::SharedPtr<TLayoutEntityIList> GetChildEntities(LayoutEntityType type, bool isDeep);

    protected:
        LayoutEntity();
        System::SharedPtr<System::Object> GetLayoutObject();
        void SetLayoutObject(System::SharedPtr<System::Object> value);
        System::SharedPtr<LayoutEntity> AddChildEntity(System::SharedPtr<LayoutEnumerator> it);
        virtual void SetParentNode(System::SharedPtr<Node> value);
        template <typename T> System::SharedPtr<System::Collections::Generic::IList<T>> GetChildNodes()
        {
            typedef LayoutEntity BaseT_LayoutEntity;
            assert_is_base_of(BaseT_LayoutEntity, T);
            assert_is_constructable(T);
            T obj = System::MakeObject<T>();
            System::SharedPtr<System::Collections::Generic::IList<T>> childList = System::MakeObject<System::Collections::Generic::List<T>>();
            for (System::SharedPtr<LayoutEntity> entity : System::IterateOver(mChildEntities))
            {
                if (System::ObjectExt::GetType(entity) == System::ObjectExt::GetType(obj))
                {
                    childList->Add(System::StaticCast<typename T::Pointee_>(entity));
                }
            }
            return childList;
        }
        System::Object::shared_members_type GetSharedMembers() override;

        System::String mKind;
        int32_t mPageIndex;
        System::SharedPtr<Node> mParentNode;
        System::Drawing::RectangleF mRectangle;
        LayoutEntityType mType;
        System::SharedPtr<LayoutEntity> mParent;
        System::SharedPtr<TLayoutEntityIList> mChildEntities;

    private:
        System::SharedPtr<System::Object> mLayoutObject;
    };

    RTTI_INFO_IMPL_HASH(3019693939u, LayoutEntity, ThisTypeBaseTypesInfo);

    int32_t LayoutEntity::GetPageIndex()
    {
        return mPageIndex;
    }

    System::Drawing::RectangleF LayoutEntity::GetRectangle()
    {
        return mRectangle;
    }

    LayoutEntityType LayoutEntity::GetLayoutEntityType()
    {
        return mType;
    }

    System::String LayoutEntity::GetText()
    {
        System::SharedPtr<System::Text::StringBuilder> builder = System::MakeObject<System::Text::StringBuilder>();
        for (System::SharedPtr<LayoutEntity> entity : System::IterateOver(mChildEntities))
        {
            builder->Append(entity->GetText());
        }

        return builder->ToString();
    }

    System::SharedPtr<LayoutEntity> LayoutEntity::GetParent()
    {
        return mParent;
    }

    System::SharedPtr<Node> LayoutEntity::GetParentNode()
    {
        return mParentNode;
    }

    System::SharedPtr<System::Object> LayoutEntity::GetLayoutObject()
    {
        return mLayoutObject;
    }

    void LayoutEntity::SetLayoutObject(System::SharedPtr<System::Object> value)
    {
        mLayoutObject = value;
    }

    LayoutEntity::LayoutEntity() : mPageIndex(0), mType(LayoutEntityType::None), mChildEntities(System::MakeObject<TLayoutEntityList>())
    {
    }

    void LayoutEntity::SetParentNode(System::SharedPtr<Node> value)
    {
        mParentNode = value;
    }

    System::SharedPtr<LayoutEntity> LayoutEntity::AddChildEntity(System::SharedPtr<LayoutEnumerator> it)
    {
        System::SharedPtr<LayoutEntity> child = CreateLayoutEntityFromType(it);
        // init child fields
        child->mKind = it->get_Kind();
        child->mPageIndex = it->get_PageIndex();
        child->mRectangle = it->get_Rectangle();
        child->mType = it->get_Type();
        child->SetLayoutObject(it->get_Current());
        child->mParent = System::MakeSharedPtr(this);
        mChildEntities->Add(child);
        return child;
    }

    System::SharedPtr<TLayoutEntityIList> LayoutEntity::GetChildEntities(LayoutEntityType type, bool isDeep)
    {
        System::SharedPtr<TLayoutEntityList> childList = System::MakeObject<TLayoutEntityList>();

        for (System::SharedPtr<LayoutEntity> entity : System::IterateOver(mChildEntities))
        {
            if ((entity->GetLayoutEntityType() & type) == entity->GetLayoutEntityType())
            {
                childList->Add(entity);
            }

            if (isDeep)
            {
                childList->AddRange(entity->GetChildEntities(type, true));
            }
        }

        return childList;
    }

    System::Object::shared_members_type LayoutEntity::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("LayoutEntity::mParentNode", this->mParentNode);
        result.Add("LayoutEntity::mRectangle", this->mRectangle);
        result.Add("LayoutEntity::mType", this->mType);
        result.Add("LayoutEntity::mParent", this->mParent);
        result.Add("LayoutEntity::mChildEntities", this->mChildEntities);
        return result;
    }

    /// <summary>
    /// Represents an entity that contains lines and rows.
    /// </summary>
    class StoryLayoutEntity : public LayoutEntity
    {
        typedef StoryLayoutEntity ThisType;
        typedef LayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::SharedPtr<TRenderedLineIList> GetLines();
        System::SharedPtr<TRenderedRowIList> GetRows();
    };

    RTTI_INFO_IMPL_HASH(3178778940u, StoryLayoutEntity, ThisTypeBaseTypesInfo);

    System::SharedPtr<TRenderedLineIList> StoryLayoutEntity::GetLines()
    {
        return GetChildNodes<System::SharedPtr<RenderedLine>>();
    }

    System::SharedPtr<TRenderedRowIList> StoryLayoutEntity::GetRows()
    {
        return GetChildNodes<System::SharedPtr<RenderedRow>>();
    }

    /// <summary>
    /// Represents line of characters of text and inline objects.
    /// </summary>
    class RenderedLine : public LayoutEntity
    {
        typedef RenderedLine ThisType;
        typedef LayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::String GetText() override;
        System::SharedPtr<Paragraph> GetParagraph();
        System::SharedPtr<TRenderedSpanIList> GetSpans();
    };

    RTTI_INFO_IMPL_HASH(3891846319u, RenderedLine, ThisTypeBaseTypesInfo);

    System::String RenderedLine::GetText()
    {
        return LayoutEntity::GetText() + System::Environment::get_NewLine();
    }

    System::SharedPtr<Paragraph> RenderedLine::GetParagraph()
    {
        return System::DynamicCast<Paragraph>(GetParentNode());
    }

    System::SharedPtr<TRenderedSpanIList> RenderedLine::GetSpans()
    {
        return GetChildNodes<System::SharedPtr<RenderedSpan>>();
    }

    /// <summary>
    /// Represents one or more characters in a line.
    /// This include special characters like field start/end markers, bookmarks, shapes and comments.
    /// </summary>
    class RenderedSpan : public LayoutEntity
    {
        typedef RenderedSpan ThisType;
        typedef LayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        RenderedSpan();
        RenderedSpan(System::String text);
        System::String GetKind();
        System::String GetText() override;

    private:
        System::String mText;
    };

    RTTI_INFO_IMPL_HASH(822999877u, RenderedSpan, ThisTypeBaseTypesInfo);

    System::String RenderedSpan::GetKind()
    {
        return mKind;
    }

    System::String RenderedSpan::GetText()
    {
        return mText;
    }

    RenderedSpan::RenderedSpan()
    {
    }


    RenderedSpan::RenderedSpan(System::String text)
    {
        mText = text != nullptr ? text : System::String::Empty;
    }

    /// <summary>
    /// Represents the header/footer content on a page.
    /// </summary>
    class RenderedHeaderFooter : public StoryLayoutEntity
    {
        typedef RenderedHeaderFooter ThisType;
        typedef StoryLayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::String GetKind();
    };

    RTTI_INFO_IMPL_HASH(1631085443u, RenderedHeaderFooter, ThisTypeBaseTypesInfo);

    System::String RenderedHeaderFooter::GetKind()
    {
        return mKind;
    }

    /// <summary>
    /// Represents a table cell.
    /// </summary>
    class RenderedCell : public StoryLayoutEntity
    {
        typedef RenderedCell ThisType;
        typedef StoryLayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::SharedPtr<Cell> GetCell();
        System::SharedPtr<Node> GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(1128351613u, RenderedCell, ThisTypeBaseTypesInfo);

    System::SharedPtr<Cell> RenderedCell::GetCell()
    {
        return System::DynamicCast<Cell>(GetParentNode());
    }

    System::SharedPtr<Node> RenderedCell::GetParentNode()
    {
        if (GetFirst(GetLines()) == nullptr)
        {
            return nullptr;
        }
        else
        {
            return GetFirst(GetLines())->GetParagraph() != nullptr ? GetFirst(GetLines())->GetParagraph()->GetAncestor(NodeType::Cell) : nullptr;
        }
    }

    /// <summary>
    /// Represents a table row.
    /// </summary>
    class RenderedRow : public LayoutEntity
    {
        typedef RenderedRow ThisType;
        typedef LayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::SharedPtr<TRenderedCellIList> GetCells();
        System::SharedPtr<Row> GetRow();
        System::SharedPtr<Table> GetTable();
        System::SharedPtr<Node> GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(1922289599u, RenderedRow, ThisTypeBaseTypesInfo);

    System::SharedPtr<TRenderedCellIList> RenderedRow::GetCells()
    {
        return GetChildNodes<System::SharedPtr<RenderedCell>>();
    }

    System::SharedPtr<Row> RenderedRow::GetRow()
    {
        return System::DynamicCast<Row>(GetParentNode());
    }

    System::SharedPtr<Table> RenderedRow::GetTable()
    {
        return GetRow() != nullptr ? GetRow()->get_ParentTable() : nullptr;
    }

    System::SharedPtr<Node> RenderedRow::GetParentNode()
    {
        System::SharedPtr<Paragraph> para = GetFirst(GetFirst(GetCells())->GetLines()) != nullptr ? GetFirst(GetFirst(GetCells())->GetLines())->GetParagraph() : nullptr;
        return para != nullptr ? para->GetAncestor(NodeType::Row) : para;
    }

    /// <summary>
    /// Represents a column of text on a page.
    /// </summary>
    class RenderedColumn : public StoryLayoutEntity
    {
        typedef RenderedColumn ThisType;
        typedef StoryLayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::SharedPtr<TRenderedFootnoteIList> GetFootnotes();
        System::SharedPtr<TRenderedEndnoteIList> GetEndnotes();
        System::SharedPtr<TRenderedNoteSeparatorIList> GetNoteSeparators();
        System::SharedPtr<Body> GetBody();
        System::SharedPtr<Node> GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(2461185489u, RenderedColumn, ThisTypeBaseTypesInfo);

    System::SharedPtr<TRenderedFootnoteIList> RenderedColumn::GetFootnotes()
    {
        return GetChildNodes<System::SharedPtr<RenderedFootnote>>();
    }

    System::SharedPtr<TRenderedEndnoteIList> RenderedColumn::GetEndnotes()
    {
        return GetChildNodes<System::SharedPtr<RenderedEndnote>>();
    }

    System::SharedPtr<TRenderedNoteSeparatorIList> RenderedColumn::GetNoteSeparators()
    {
        return GetChildNodes<System::SharedPtr<RenderedNoteSeparator>>();
    }

    System::SharedPtr<Body> RenderedColumn::GetBody()
    {
        return System::DynamicCast<Body>(GetParentNode());
    }

    System::SharedPtr<Node> RenderedColumn::GetParentNode()
    {
        return GetFirst(GetChildEntities(LayoutEntityType::Line, true))->GetParentNode()->GetAncestor(NodeType::Body);
    }

    /// <summary>
    /// Represents placeholder for footnote content.
    /// </summary>
    class RenderedFootnote : public StoryLayoutEntity
    {
        typedef RenderedFootnote ThisType;
        typedef StoryLayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::SharedPtr<Footnote> GetFootnote();
        System::SharedPtr<Node> GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(4175017083u, RenderedFootnote, ThisTypeBaseTypesInfo);

    System::SharedPtr<Footnote> RenderedFootnote::GetFootnote()
    {
        return System::DynamicCast<Footnote>(GetParentNode());
    }

    System::SharedPtr<Node> RenderedFootnote::GetParentNode()
    {
        return GetFirst(GetChildEntities(LayoutEntityType::Line, true))->GetParentNode()->GetAncestor(NodeType::Footnote);
    }

    /// <summary>
    /// Represents placeholder for endnote content.
    /// </summary>
    class RenderedEndnote : public StoryLayoutEntity
    {
        typedef RenderedEndnote ThisType;
        typedef StoryLayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::SharedPtr<Footnote> GetEndnote();
        System::SharedPtr<Node> GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(31763922u, RenderedEndnote, ThisTypeBaseTypesInfo);

    System::SharedPtr<Footnote> RenderedEndnote::GetEndnote()
    {
        return System::DynamicCast<Footnote>(GetParentNode());
    }

    System::SharedPtr<Node> RenderedEndnote::GetParentNode()
    {
        return GetFirst(GetChildEntities(LayoutEntityType::Line, true))->GetParentNode()->GetAncestor(NodeType::Footnote);
    }

    /// <summary>
    /// Represents text area inside of a shape.
    /// </summary>
    class RenderedTextBox : public StoryLayoutEntity
    {
        typedef RenderedTextBox ThisType;
        typedef StoryLayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::SharedPtr<Node> GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(374067683u, RenderedTextBox, ThisTypeBaseTypesInfo);

    System::SharedPtr<Node> RenderedTextBox::GetParentNode()
    {
        System::SharedPtr<TLayoutEntityIList> lines = GetChildEntities(LayoutEntityType::Line, true);
        System::SharedPtr<Node> shape = GetFirst(lines)->GetParentNode()->GetAncestor(NodeType::Shape);

        if (shape != nullptr)
        {
            return shape;
        }
        else
        {
            return GetFirst(lines)->GetParentNode()->GetAncestor(NodeType::Shape);
        }
    }

    /// <summary>
    /// Represents placeholder for comment content.
    /// </summary>
    class RenderedComment : public StoryLayoutEntity
    {
        typedef RenderedComment ThisType;
        typedef StoryLayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::SharedPtr<Comment> GetComment();
        System::SharedPtr<Node> GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(2609978340u, RenderedComment, ThisTypeBaseTypesInfo);

    System::SharedPtr<Aspose::Words::Comment> RenderedComment::GetComment()
    {
        return System::DynamicCast<Comment>(GetParentNode());
    }

    System::SharedPtr<Node> RenderedComment::GetParentNode()
    {
        return GetFirst(GetChildEntities(LayoutEntityType::Line, true))->GetParentNode()->GetAncestor(NodeType::Comment);
    }

    /// <summary>
    /// Represents footnote/endnote separator.
    /// </summary>
    class RenderedNoteSeparator : public StoryLayoutEntity
    {
        typedef RenderedNoteSeparator ThisType;
        typedef StoryLayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::SharedPtr<Footnote> GetFootnote();
        System::SharedPtr<Node> GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(425844120u, RenderedNoteSeparator, ThisTypeBaseTypesInfo);

    System::SharedPtr<Footnote> RenderedNoteSeparator::GetFootnote()
    {
        return System::DynamicCast<Footnote>(GetParentNode());
    }

    System::SharedPtr<Node> RenderedNoteSeparator::GetParentNode()
    {
        return GetFirst(GetChildEntities(LayoutEntityType::Line, true))->GetParentNode()->GetAncestor(NodeType::Footnote);
    }

    /// <summary>
    /// Represents page of a document.
    /// </summary>
    class RenderedPage : public LayoutEntity
    {
        typedef RenderedPage ThisType;
        typedef LayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        System::SharedPtr<TRenderedColumnIList> GetColumns();
        System::SharedPtr<TRenderedHeaderFooterIList> GetHeaderFooters();
        System::SharedPtr<TRenderedCommentIList> GetComments();
        System::SharedPtr<Section> GetSection();
        System::SharedPtr<Node> GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(2652676586u, RenderedPage, ThisTypeBaseTypesInfo);

    System::SharedPtr<TRenderedColumnIList> RenderedPage::GetColumns()
    {
        return GetChildNodes<System::SharedPtr<RenderedColumn>>();
    }

    System::SharedPtr<TRenderedHeaderFooterIList> RenderedPage::GetHeaderFooters()
    {
        return GetChildNodes<System::SharedPtr<RenderedHeaderFooter>>();
    }

    System::SharedPtr<TRenderedCommentIList> RenderedPage::GetComments()
    {
        return GetChildNodes<System::SharedPtr<RenderedComment>>();
    }

    System::SharedPtr<Section> RenderedPage::GetSection()
    {
        return System::DynamicCast<Aspose::Words::Section>(GetParentNode());
    }

    System::SharedPtr<Node> RenderedPage::GetParentNode()
    {
        return GetFirst(GetFirst(GetColumns())->GetChildEntities(LayoutEntityType::Line, true))->GetParentNode()->GetAncestor(NodeType::Section);
    }

    /// <summary>
    /// Provides an API wrapper for the LayoutEnumerator class to access the page layout entities of a document presented in an object model like design.
    /// </summary>
    class RenderedDocument : public LayoutEntity
    {
        typedef RenderedDocument ThisType;
        typedef LayoutEntity BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        RenderedDocument(System::SharedPtr<Document> doc);
        System::SharedPtr<TRenderedPageIList> GetPages();
        System::SharedPtr<TLayoutEntityIList> GetLayoutEntitiesOfNode(System::SharedPtr<Node> node);

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        void ProcessLayoutElements(System::SharedPtr<LayoutEntity> current);
        void CollectLinesAndAddToMarkers();
        void CollectLinesOfMarkersCore(LayoutEntityType type);
        void LinkLayoutMarkersToNodes(System::SharedPtr<Document> doc);

        System::SharedPtr<LayoutCollector> mLayoutCollector;
        System::SharedPtr<LayoutEnumerator> mEnumerator;
        static System::SharedPtr<LayoutNodeLookupIDict> mLayoutToNodeLookup;
    };

    RTTI_INFO_IMPL_HASH(3105141718u, RenderedDocument, ThisTypeBaseTypesInfo);

    System::SharedPtr<LayoutNodeLookupIDict> RenderedDocument::mLayoutToNodeLookup = System::MakeObject<LayoutNodeLookupDict>();

    System::SharedPtr<TRenderedPageIList> RenderedDocument::GetPages()
    {
        return GetChildNodes<System::SharedPtr<RenderedPage>>();
    }

    RenderedDocument::RenderedDocument(System::SharedPtr<Document> doc)
    {
        //Self reference protector
        System::Details::ThisProtector __local_self_ref(this);

        mLayoutCollector = System::MakeObject<LayoutCollector>(doc);
        mEnumerator = System::MakeObject<LayoutEnumerator>(doc);
        ProcessLayoutElements(System::MakeSharedPtr(this));
        LinkLayoutMarkersToNodes(doc);
        CollectLinesAndAddToMarkers();
    }

    System::SharedPtr<TLayoutEntityIList> RenderedDocument::GetLayoutEntitiesOfNode(System::SharedPtr<Node> node)
    {
        if (!System::ObjectExt::Equals(mLayoutCollector->get_Document(), node->get_Document()))
        {
            throw System::ArgumentException(u"Node does not belong to the same document which was rendered.");
        }

        if (node->get_NodeType() == NodeType::Document)
        {
            return mChildEntities;
        }

        System::SharedPtr<TLayoutEntityIList> entities = System::MakeObject<TLayoutEntityList>();

        // Retrieve all entities from the layout document (inversion of LayoutEntityType.None).
        for (System::SharedPtr<LayoutEntity> entity : System::IterateOver(GetChildEntities(~LayoutEntityType::None, true)))
        {
            if (entity->GetParentNode() == node)
            {
                entities->Add(entity);
            }

            // There is no table entity in rendered output so manually check if rows belong to a table node.
            if (entity->GetLayoutEntityType() == LayoutEntityType::Row)
            {
                System::SharedPtr<RenderedRow> row = System::StaticCast<RenderedRow>(entity);
                if (row->GetTable() == node)
                {
                    entities->Add(entity);
                }
            }
        }

        return entities;
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

                current = current->GetParent();
            }
        } while (mEnumerator->MoveNext());
    }

    void RenderedDocument::CollectLinesAndAddToMarkers()
    {
        CollectLinesOfMarkersCore(LayoutEntityType::Column);
        CollectLinesOfMarkersCore(Layout::LayoutEntityType::Comment);
    }

    void RenderedDocument::CollectLinesOfMarkersCore(LayoutEntityType type)
    {
        System::SharedPtr<TRenderedLineIList> collectedLines = System::MakeObject<TRenderedLineList>();
        for (System::SharedPtr<RenderedPage> page : System::IterateOver(GetPages()))
        {
            for (System::SharedPtr<LayoutEntity> story : System::IterateOver(page->GetChildEntities(type, false)))
            {
                for (System::SharedPtr<RenderedLine> line : System::IterateOver<RenderedLine>(story->GetChildEntities(LayoutEntityType::Line, true)))
                {
                    collectedLines->Add(line);
                    for (System::SharedPtr<RenderedSpan> span : System::IterateOver(line->GetSpans()))
                    {
                        if (mLayoutToNodeLookup->ContainsKey(span->GetLayoutObject()))
                        {
                            if (span->GetKind() == u"PARAGRAPH" || span->GetKind() == u"ROW" || span->GetKind() == u"CELL" || span->GetKind() == u"SECTION")
                            {
                                System::SharedPtr<Node> node = mLayoutToNodeLookup->idx_get(span->GetLayoutObject());

                                if (node->get_NodeType() == NodeType::Row)
                                {
                                    node = (System::DynamicCast<Row>(node))->get_LastCell()->get_LastParagraph();
                                }

                                for (System::SharedPtr<RenderedLine> collectedLine : System::IterateOver(collectedLines))
                                {
                                    collectedLine->SetParentNode(node);
                                }

                                collectedLines = System::MakeObject<TRenderedLineList>();
                            }
                            else
                            {
                                span->SetParentNode(mLayoutToNodeLookup->idx_get(span->GetLayoutObject()));
                            }
                        }
                    }
                }
            }
        }
    }

    void RenderedDocument::LinkLayoutMarkersToNodes(System::SharedPtr<Document> doc)
    {
        for (System::SharedPtr<Node> node : System::IterateOver(doc->GetChildNodes(NodeType::Any, true)))
        {
            System::SharedPtr<System::Object> entity = mLayoutCollector->GetEntity(node);

            if (entity != nullptr)
            {
                mLayoutToNodeLookup->Add(entity, node);
            }
        }
    }

    System::Object::shared_members_type RenderedDocument::GetSharedMembers()
    {
        auto result = LayoutEntity::GetSharedMembers();
        result.Add("RenderedDocument::mLayoutCollector", this->mLayoutCollector);
        result.Add("RenderedDocument::mEnumerator", this->mEnumerator);
        return result;
    }

    System::SharedPtr<LayoutEntity> CreateLayoutEntityFromType(System::SharedPtr<LayoutEnumerator> it)
    {
        System::SharedPtr<LayoutEntity> childEntity;
        switch (it->get_Type())
        {
        case LayoutEntityType::Cell:
            return System::MakeObject<RenderedCell>();

        case LayoutEntityType::Column:
            return System::MakeObject<RenderedColumn>();

        case LayoutEntityType::Comment:
            return System::MakeObject<RenderedComment>();

        case LayoutEntityType::Endnote:
            return System::MakeObject<RenderedEndnote>();

        case LayoutEntityType::Footnote:
            return System::MakeObject<RenderedFootnote>();

        case LayoutEntityType::HeaderFooter:
            return System::MakeObject<RenderedHeaderFooter>();

        case LayoutEntityType::Line:
            return System::MakeObject<RenderedLine>();

        case LayoutEntityType::NoteSeparator:
            return System::MakeObject<RenderedNoteSeparator>();

        case LayoutEntityType::Page:
            return System::MakeObject<RenderedPage>();

        case LayoutEntityType::Row:
            return System::MakeObject<RenderedRow>();

        case LayoutEntityType::Span:
            return System::MakeObject<RenderedSpan>(it->get_Text());

        case LayoutEntityType::TextBox:
            return System::MakeObject<RenderedTextBox>();

        default:
            throw System::InvalidOperationException(u"Unknown layout type");
        }
    }
}

void DocumentLayoutHelper()
{
    std::cout << "DocumentLayoutHelper example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_RenderingAndPrinting();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.docx");

    // This sample introduces the RenderedDocument class and other related classes which provide an API wrapper for 
    // The LayoutEnumerator. This allows you to access the layout entities of a document using a DOM style API.

    // Create a new RenderedDocument class from a Document object.
    System::SharedPtr<RenderedDocument> layoutDoc = System::MakeObject<RenderedDocument>(doc);

    // The following examples demonstrate how to use the wrapper API. 
    // This snippet returns the third line of the first page and prints the line of text to the console.
    System::SharedPtr<RenderedLine> line = layoutDoc->GetPages()->idx_get(0)->GetColumns()->idx_get(0)->GetLines()->idx_get(2);
    std::cout << "Line: " << line->GetText().ToUtf8String() << std::endl;

    // With a rendered line the original paragraph in the document object model can be returned.
    System::SharedPtr<Paragraph> para = line->GetParagraph();
    std::cout << "Paragraph text: " << para->get_Range()->get_Text().ToUtf8String() << std::endl;

    // Retrieve all the text that appears of the first page in plain text format (including headers and footers).
    System::String pageText = layoutDoc->GetPages()->idx_get(0)->GetText();
    std::cout << "Page text: " << pageText.ToUtf8String() << std::endl;

    // Loop through each page in the document and print how many lines appear on each page.
    for (System::SharedPtr<RenderedPage> page : System::IterateOver(layoutDoc->GetPages()))
    {
        System::SharedPtr<System::Collections::Generic::IList<System::SharedPtr<LayoutEntity>>> lines = page->GetChildEntities(LayoutEntityType::Line, true);
        std::cout << "Page " << page->GetPageIndex() << " has " << lines->get_Count() << " lines." << std::endl;
    }

    // This method provides a reverse lookup of layout entities for any given node (with the exception of runs and nodes in the
    // Header and footer).
    std::cout << std::endl << "The lines of the second paragraph:" << std::endl;
    for (System::SharedPtr<RenderedLine> paragraphLine : System::IterateOver<RenderedLine>(layoutDoc->GetLayoutEntitiesOfNode(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1))))
    {
        std::cout << "\"" << paragraphLine->GetText().Trim().ToUtf8String() << "\"" << std::endl;
        std::cout << paragraphLine->GetRectangle().ToString().ToUtf8String() << std::endl << std::endl;
    }

    std::cout << "Document layout helper example ran successfully." << std::endl;
    std::cout << "DocumentLayoutHelper example finished." << std::endl << std::endl;
}