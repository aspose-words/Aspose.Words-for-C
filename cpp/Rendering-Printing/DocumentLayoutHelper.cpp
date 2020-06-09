#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutCollector.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEntityType.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Tables;

namespace
{
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

    typedef System::SharedPtr<LayoutEntity> TLayoutEntityPtr;
    typedef System::SharedPtr<RenderedLine> TRenderedLinePtr;
    typedef System::SharedPtr<RenderedRow> TRenderedRowPtr;
    typedef System::SharedPtr<RenderedSpan> TRenderedSpanPtr;
    typedef System::SharedPtr<RenderedCell> TRenderedCellPtr;
    typedef System::SharedPtr<RenderedFootnote> TRenderedFootnotePtr;
    typedef System::SharedPtr<RenderedEndnote> TRenderedEndnotePtr;
    typedef System::SharedPtr<RenderedNoteSeparator> TRenderedNoteSeparatorPtr;
    typedef System::SharedPtr<RenderedColumn> TRenderedColumnPtr;
    typedef System::SharedPtr<RenderedHeaderFooter> TRenderedHeaderFooterPtr;
    typedef System::SharedPtr<RenderedComment> TRenderedCommentPtr;
    typedef System::SharedPtr<RenderedPage> TRenderedPagePtr;
    typedef System::SharedPtr<System::Object> TObjectPtr;
    typedef System::SharedPtr<Node> TNodePtr;

    TLayoutEntityPtr CreateLayoutEntityFromType(System::SharedPtr<LayoutEnumerator> it);

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
        TLayoutEntityPtr GetParent();
        virtual TNodePtr GetParentNode();
        std::vector<TLayoutEntityPtr> GetChildEntities(LayoutEntityType type, bool isDeep);

    protected:
        LayoutEntity();
        TObjectPtr GetLayoutObject();
        void SetLayoutObject(TObjectPtr value);
        TLayoutEntityPtr AddChildEntity(System::SharedPtr<LayoutEnumerator> it);
        virtual void SetParentNode(TNodePtr value);
        template <typename T> std::vector<T> GetChildNodes()
        {
            typedef LayoutEntity BaseT_LayoutEntity;
            assert_is_base_of(BaseT_LayoutEntity, T);
            assert_is_constructable(T);
            T obj = System::MakeObject<T>();
            std::vector<T> childList;
            for (TLayoutEntityPtr entity : mChildEntities)
            {
                if (System::ObjectExt::GetType(entity) == System::ObjectExt::GetType(obj))
                {
                    childList.push_back(System::StaticCast<typename T::Pointee_>(entity));
                }
            }
            return childList;
        }

        System::String mKind;
        int32_t mPageIndex;
        TNodePtr mParentNode;
        System::Drawing::RectangleF mRectangle;
        LayoutEntityType mType;
        TLayoutEntityPtr mParent;
        std::vector<TLayoutEntityPtr> mChildEntities;

    private:
        TObjectPtr mLayoutObject;
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
        for (TLayoutEntityPtr entity : mChildEntities)
        {
            builder->Append(entity->GetText());
        }

        return builder->ToString();
    }

    TLayoutEntityPtr LayoutEntity::GetParent()
    {
        return mParent;
    }

    TNodePtr LayoutEntity::GetParentNode()
    {
        return mParentNode;
    }

    TObjectPtr LayoutEntity::GetLayoutObject()
    {
        return mLayoutObject;
    }

    void LayoutEntity::SetLayoutObject(TObjectPtr value)
    {
        mLayoutObject = value;
    }

    LayoutEntity::LayoutEntity() : mPageIndex(0), mType(LayoutEntityType::None)
    {
    }

    void LayoutEntity::SetParentNode(TNodePtr value)
    {
        mParentNode = value;
    }

    TLayoutEntityPtr LayoutEntity::AddChildEntity(System::SharedPtr<LayoutEnumerator> it)
    {
        TLayoutEntityPtr child = CreateLayoutEntityFromType(it);
        // init child fields
        child->mKind = it->get_Kind();
        child->mPageIndex = it->get_PageIndex();
        child->mRectangle = it->get_Rectangle();
        child->mType = it->get_Type();
        child->SetLayoutObject(it->get_Current());
        child->mParent = System::MakeSharedPtr(this);
        mChildEntities.push_back(child);
        return child;
    }

    std::vector<TLayoutEntityPtr> LayoutEntity::GetChildEntities(LayoutEntityType type, bool isDeep)
    {
        std::vector<TLayoutEntityPtr> childList;

        for (TLayoutEntityPtr entity : mChildEntities)
        {
            if ((entity->GetLayoutEntityType() & type) == entity->GetLayoutEntityType())
            {
                childList.push_back(entity);
            }

            if (isDeep)
            {
                std::vector<TLayoutEntityPtr> children(entity->GetChildEntities(type, true));
                childList.insert(childList.end(), children.begin(), children.end());
            }
        }

        return childList;
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
        std::vector<TRenderedLinePtr> GetLines();
        std::vector<TRenderedRowPtr> GetRows();
    };

    RTTI_INFO_IMPL_HASH(3178778940u, StoryLayoutEntity, ThisTypeBaseTypesInfo);

    std::vector<TRenderedLinePtr> StoryLayoutEntity::GetLines()
    {
        return GetChildNodes<TRenderedLinePtr>();
    }

    std::vector<TRenderedRowPtr> StoryLayoutEntity::GetRows()
    {
        return GetChildNodes<TRenderedRowPtr>();
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
        std::vector<TRenderedSpanPtr> GetSpans();
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

    std::vector<TRenderedSpanPtr> RenderedLine::GetSpans()
    {
        return GetChildNodes<TRenderedSpanPtr>();
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
        RenderedSpan(System::String const &text);
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


    RenderedSpan::RenderedSpan(System::String const &text)
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
        TNodePtr GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(1128351613u, RenderedCell, ThisTypeBaseTypesInfo);

    System::SharedPtr<Cell> RenderedCell::GetCell()
    {
        return System::DynamicCast<Cell>(GetParentNode());
    }

    TNodePtr RenderedCell::GetParentNode()
    {
        std::vector<TRenderedLinePtr> lines(GetLines());
        if (lines.empty())
        {
            return nullptr;
        }
        else
        {
            return lines.at(0)->GetParagraph() != nullptr ? lines.at(0)->GetParagraph()->GetAncestor(NodeType::Cell) : nullptr;
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
        std::vector<TRenderedCellPtr> GetCells();
        System::SharedPtr<Row> GetRow();
        System::SharedPtr<Table> GetTable();
        TNodePtr GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(1922289599u, RenderedRow, ThisTypeBaseTypesInfo);

    std::vector<TRenderedCellPtr> RenderedRow::GetCells()
    {
        return GetChildNodes<TRenderedCellPtr>();
    }

    System::SharedPtr<Row> RenderedRow::GetRow()
    {
        return System::DynamicCast<Row>(GetParentNode());
    }

    System::SharedPtr<Table> RenderedRow::GetTable()
    {
        return GetRow() != nullptr ? GetRow()->get_ParentTable() : nullptr;
    }

    TNodePtr RenderedRow::GetParentNode()
    {
        std::vector<TRenderedCellPtr> cells(GetCells());
        if (cells.empty())
            return nullptr;
        std::vector<TRenderedLinePtr> lines(cells.at(0)->GetLines());
        System::SharedPtr<Paragraph> para = lines.empty() ? nullptr : lines.at(0)->GetParagraph();
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
        std::vector<TRenderedFootnotePtr> GetFootnotes();
        std::vector<TRenderedEndnotePtr> GetEndnotes();
        std::vector<TRenderedNoteSeparatorPtr> GetNoteSeparators();
        System::SharedPtr<Body> GetBody();
        TNodePtr GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(2461185489u, RenderedColumn, ThisTypeBaseTypesInfo);

    std::vector<TRenderedFootnotePtr> RenderedColumn::GetFootnotes()
    {
        return GetChildNodes<TRenderedFootnotePtr>();
    }

    std::vector<TRenderedEndnotePtr> RenderedColumn::GetEndnotes()
    {
        return GetChildNodes<TRenderedEndnotePtr>();
    }

    std::vector<TRenderedNoteSeparatorPtr> RenderedColumn::GetNoteSeparators()
    {
        return GetChildNodes<TRenderedNoteSeparatorPtr>();
    }

    System::SharedPtr<Body> RenderedColumn::GetBody()
    {
        return System::DynamicCast<Body>(GetParentNode());
    }

    TNodePtr RenderedColumn::GetParentNode()
    {
        return GetChildEntities(LayoutEntityType::Line, true).at(0)->GetParentNode()->GetAncestor(NodeType::Body);
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
        TNodePtr GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(4175017083u, RenderedFootnote, ThisTypeBaseTypesInfo);

    System::SharedPtr<Footnote> RenderedFootnote::GetFootnote()
    {
        return System::DynamicCast<Footnote>(GetParentNode());
    }

    TNodePtr RenderedFootnote::GetParentNode()
    {
        return GetChildEntities(LayoutEntityType::Line, true).at(0)->GetParentNode()->GetAncestor(NodeType::Footnote);
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
        TNodePtr GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(31763922u, RenderedEndnote, ThisTypeBaseTypesInfo);

    System::SharedPtr<Footnote> RenderedEndnote::GetEndnote()
    {
        return System::DynamicCast<Footnote>(GetParentNode());
    }

    TNodePtr RenderedEndnote::GetParentNode()
    {
        return GetChildEntities(LayoutEntityType::Line, true).at(0)->GetParentNode()->GetAncestor(NodeType::Footnote);
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
        TNodePtr GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(374067683u, RenderedTextBox, ThisTypeBaseTypesInfo);

    TNodePtr RenderedTextBox::GetParentNode()
    {
        std::vector<TLayoutEntityPtr> lines(GetChildEntities(LayoutEntityType::Line, true));
        TNodePtr shape = lines.at(0)->GetParentNode()->GetAncestor(NodeType::Shape);
        return shape;
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
        TNodePtr GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(2609978340u, RenderedComment, ThisTypeBaseTypesInfo);

    System::SharedPtr<Comment> RenderedComment::GetComment()
    {
        return System::DynamicCast<Comment>(GetParentNode());
    }

    TNodePtr RenderedComment::GetParentNode()
    {
        return GetChildEntities(LayoutEntityType::Line, true).at(0)->GetParentNode()->GetAncestor(NodeType::Comment);
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
        TNodePtr GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(425844120u, RenderedNoteSeparator, ThisTypeBaseTypesInfo);

    System::SharedPtr<Footnote> RenderedNoteSeparator::GetFootnote()
    {
        return System::DynamicCast<Footnote>(GetParentNode());
    }

    TNodePtr RenderedNoteSeparator::GetParentNode()
    {
        return GetChildEntities(LayoutEntityType::Line, true).at(0)->GetParentNode()->GetAncestor(NodeType::Footnote);
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
        std::vector<TRenderedColumnPtr> GetColumns();
        std::vector<TRenderedHeaderFooterPtr> GetHeaderFooters();
        std::vector<TRenderedCommentPtr> GetComments();
        System::SharedPtr<Section> GetSection();
        TNodePtr GetParentNode() override;
    };

    RTTI_INFO_IMPL_HASH(2652676586u, RenderedPage, ThisTypeBaseTypesInfo);

    std::vector<TRenderedColumnPtr> RenderedPage::GetColumns()
    {
        return GetChildNodes<TRenderedColumnPtr>();
    }

    std::vector<TRenderedHeaderFooterPtr> RenderedPage::GetHeaderFooters()
    {
        return GetChildNodes<TRenderedHeaderFooterPtr>();
    }

    std::vector<TRenderedCommentPtr> RenderedPage::GetComments()
    {
        return GetChildNodes<TRenderedCommentPtr>();
    }

    System::SharedPtr<Section> RenderedPage::GetSection()
    {
        return System::DynamicCast<Section>(GetParentNode());
    }

    TNodePtr RenderedPage::GetParentNode()
    {
        return GetColumns().at(0)->GetChildEntities(LayoutEntityType::Line, true).at(0)->GetParentNode()->GetAncestor(NodeType::Section);
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
        std::vector<TRenderedPagePtr> GetPages();
        std::vector<TLayoutEntityPtr> GetLayoutEntitiesOfNode(TNodePtr node);
    private:
        void ProcessLayoutElements(TLayoutEntityPtr current);
        void CollectLinesAndAddToMarkers();
        void CollectLinesOfMarkersCore(LayoutEntityType type);
        void LinkLayoutMarkersToNodes(System::SharedPtr<Document> doc);

        System::SharedPtr<LayoutCollector> mLayoutCollector;
        System::SharedPtr<LayoutEnumerator> mEnumerator;
        static std::unordered_map<TObjectPtr, TNodePtr> mLayoutToNodeLookup;
    };

    RTTI_INFO_IMPL_HASH(3105141718u, RenderedDocument, ThisTypeBaseTypesInfo);

    std::unordered_map<TObjectPtr, TNodePtr> RenderedDocument::mLayoutToNodeLookup;

    std::vector<TRenderedPagePtr> RenderedDocument::GetPages()
    {
        return GetChildNodes<TRenderedPagePtr>();
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

    std::vector<TLayoutEntityPtr> RenderedDocument::GetLayoutEntitiesOfNode(TNodePtr node)
    {
        if (!System::ObjectExt::Equals(mLayoutCollector->get_Document(), node->get_Document()))
        {
            throw System::ArgumentException(u"Node does not belong to the same document which was rendered.");
        }

        if (node->get_NodeType() == NodeType::Document)
        {
            return mChildEntities;
        }

        std::vector<TLayoutEntityPtr> entities;

        // Retrieve all entities from the layout document (inversion of LayoutEntityType.None).
        for (TLayoutEntityPtr entity : GetChildEntities(~LayoutEntityType::None, true))
        {
            if (entity->GetParentNode() == node)
            {
                entities.push_back(entity);
            }

            // There is no table entity in rendered output so manually check if rows belong to a table node.
            if (entity->GetLayoutEntityType() == LayoutEntityType::Row)
            {
                TRenderedRowPtr row = System::StaticCast<RenderedRow>(entity);
                if (row->GetTable() == node)
                {
                    entities.push_back(entity);
                }
            }
        }

        return entities;
    }

    void RenderedDocument::ProcessLayoutElements(TLayoutEntityPtr current)
    {
        do
        {
            TLayoutEntityPtr child = current->AddChildEntity(mEnumerator);

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
        std::vector<TRenderedLinePtr> collectedLines;
        for (TRenderedPagePtr page : GetPages())
        {
            for (TLayoutEntityPtr story : page->GetChildEntities(type, false))
            {
                for (TLayoutEntityPtr lineEntity : story->GetChildEntities(LayoutEntityType::Line, true))
                {
                    TRenderedLinePtr line = System::DynamicCast<RenderedLine>(lineEntity);
                    collectedLines.push_back(line);
                    for (TRenderedSpanPtr span : line->GetSpans())
                    {
                        auto iterator = mLayoutToNodeLookup.find(span->GetLayoutObject());
                        if (iterator != mLayoutToNodeLookup.end())
                        {
                            if (span->GetKind() == u"PARAGRAPH" || span->GetKind() == u"ROW" || span->GetKind() == u"CELL" || span->GetKind() == u"SECTION")
                            {
                                TNodePtr node = iterator->second;

                                if (node->get_NodeType() == NodeType::Row)
                                {
                                    node = (System::DynamicCast<Row>(node))->get_LastCell()->get_LastParagraph();
                                }

                                for (TRenderedLinePtr collectedLine : collectedLines)
                                {
                                    collectedLine->SetParentNode(node);
                                }

                                collectedLines.clear();
                            }
                            else
                            {
                                span->SetParentNode(iterator->second);
                            }
                        }
                    }
                }
            }
        }
    }

    void RenderedDocument::LinkLayoutMarkersToNodes(System::SharedPtr<Document> doc)
    {
        for (TNodePtr node : System::IterateOver(doc->GetChildNodes(NodeType::Any, true)))
        {
            TObjectPtr entity = mLayoutCollector->GetEntity(node);

            if (entity != nullptr)
            {
                mLayoutToNodeLookup.insert({entity, node});
            }
        }
    }

    TLayoutEntityPtr CreateLayoutEntityFromType(System::SharedPtr<LayoutEnumerator> it)
    {
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
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.docx");

    // This sample introduces the RenderedDocument class and other related classes which provide an API wrapper for 
    // The LayoutEnumerator. This allows you to access the layout entities of a document using a DOM style API.

    // Create a new RenderedDocument class from a Document object.
    System::SharedPtr<RenderedDocument> layoutDoc = System::MakeObject<RenderedDocument>(doc);

    // The following examples demonstrate how to use the wrapper API. 
    // This snippet returns the third line of the first page and prints the line of text to the console.
    TRenderedLinePtr line = layoutDoc->GetPages().at(0)->GetColumns().at(0)->GetLines().at(2);
    std::cout << "Line: " << line->GetText().ToUtf8String() << std::endl;

    // With a rendered line the original paragraph in the document object model can be returned.
    System::SharedPtr<Paragraph> para = line->GetParagraph();
    std::cout << "Paragraph text: " << para->get_Range()->get_Text().ToUtf8String() << std::endl;

    // Retrieve all the text that appears of the first page in plain text format (including headers and footers).
    System::String pageText = layoutDoc->GetPages().at(0)->GetText();
    std::cout << "Page text: " << pageText.ToUtf8String() << std::endl;

    // Loop through each page in the document and print how many lines appear on each page.
    for (TRenderedPagePtr page : layoutDoc->GetPages())
    {
        std::vector<TLayoutEntityPtr> lines(page->GetChildEntities(LayoutEntityType::Line, true));
        std::cout << "Page " << page->GetPageIndex() << " has " << lines.size() << " lines." << std::endl;
    }

    // This method provides a reverse lookup of layout entities for any given node (with the exception of runs and nodes in the
    // Header and footer).
    std::cout << std::endl << "The lines of the second paragraph:" << std::endl;
    for (TLayoutEntityPtr paragraphLine : layoutDoc->GetLayoutEntitiesOfNode(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)))
    {
        std::cout << "\"" << paragraphLine->GetText().Trim().ToUtf8String() << "\"" << std::endl;
        std::cout << paragraphLine->GetRectangle().ToString().ToUtf8String() << std::endl << std::endl;
    }

    std::cout << "Document layout helper example ran successfully." << std::endl;
    std::cout << "DocumentLayoutHelper example finished." << std::endl << std::endl;
}