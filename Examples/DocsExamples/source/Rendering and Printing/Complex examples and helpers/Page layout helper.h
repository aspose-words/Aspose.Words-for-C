#pragma once

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Comment.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Layout/LayoutCollector.h>
#include <Aspose.Words.Cpp/Layout/LayoutEntityType.h>
#include <Aspose.Words.Cpp/Layout/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/Notes/Footnote.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <drawing/rectangle_f.h>
#include <system/collections/dictionary.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ienumerator.h>
#include <system/collections/list.h>
#include <system/constraints.h>
#include <system/default.h>
#include <system/details/pointer_collection_helpers.h>
#include <system/enum_helpers.h>
#include <system/func.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/type_info.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

namespace DocsExamples { namespace Complex_examples_and_helpers {
class RenderedNoteSeparator;
class RenderedEndnote;
class RenderedFootnote;
class RenderedCell;
class RenderedComment;
class RenderedColumn;
class RenderedSpan;
class RenderedRow;
class RenderedLine;
class RenderedPage;
template <typename> class LayoutCollection;
class RenderedDocument;
}} // namespace DocsExamples::Complex_examples_and_helpers

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Complex_examples_and_helpers {

class DocumentLayoutHelper : public DocsExamplesBase
{
public:
    void WrapperToAccessLayoutEntities();
};

/// <summary>
/// Provides the base class for rendered elements of a document.
/// </summary>
class LayoutEntity : public System::Object
{
    friend class RenderedDocument;

public:
    /// <summary>
    /// Gets the 1-based index of a page which contains the rendered entity.
    /// </summary>
    int get_PageIndex();
    /// <summary>
    /// Returns bounding rectangle of the entity relative to the page top left corner (in points).
    /// </summary>
    System::Drawing::RectangleF get_Rectangle();
    /// <summary>
    /// Gets the type of this layout entity.
    /// </summary>
    LayoutEntityType get_Type();
    /// <summary>
    /// Exports the contents of the entity into a string in plain text format.
    /// </summary>
    virtual String get_Text();
    /// <summary>
    /// Gets the immediate parent of this entity.
    /// </summary>
    SharedPtr<LayoutEntity> get_Parent();
    /// <summary>
    /// Returns the node that corresponds to this layout entity.
    /// </summary>
    /// <remarks>This property may return null for spans that originate
    /// from Run nodes or nodes inside the header or footer.</remarks>
    virtual SharedPtr<Node> get_ParentNode();

    /// <summary>
    /// Returns a collection of child entities which match the specified type.
    /// </summary>
    /// <param name="type">Specifies the type of entities to select.</param>
    /// <param name="isDeep">True to select from all child entities recursively.
    /// False to select only among immediate children</param>
    SharedPtr<LayoutCollection<SharedPtr<LayoutEntity>>> GetChildEntities(LayoutEntityType type, bool isDeep);

    LayoutEntity();

protected:
    /// <summary>
    /// Reserved for internal use.
    /// </summary>
    SharedPtr<System::Object> get_LayoutObject();
    /// <summary>
    /// Reserved for internal use.
    /// </summary>
    void set_LayoutObject(SharedPtr<System::Object> value);

    String mKind;
    int mPageIndex;
    SharedPtr<Node> mParentNode;
    System::Drawing::RectangleF mRectangle;
    LayoutEntityType mType;
    SharedPtr<LayoutEntity> mParent;
    SharedPtr<System::Collections::Generic::List<SharedPtr<LayoutEntity>>> mChildEntities;

    /// <summary>
    /// Internal method separate from ParentNode property to make code autoportable to VB.NET.
    /// </summary>
    virtual void SetParentNode(SharedPtr<Node> value);
    /// <summary>
    /// Reserved for internal use.
    /// </summary>
    SharedPtr<LayoutEntity> AddChildEntity(SharedPtr<LayoutEnumerator> it);
    template <typename T> SharedPtr<LayoutCollection<T>> GetChildNodes()
    {
        typedef LayoutEntity BaseT_LayoutEntity;
        assert_is_base_of(BaseT_LayoutEntity, T);
        assert_is_constructable(T);

        auto childList = MakeObject<System::Collections::Generic::List<T>>();
        for (auto entity : mChildEntities)
        {
            auto obj = dynamic_cast<typename T::Pointee_*>(entity.get());
            if (obj != nullptr)
                childList->Add(obj);
        }

        return LayoutCollection<T>::MakeObject(childList);
    }

    virtual ~LayoutEntity();

private:
    SharedPtr<System::Object> pr_LayoutObject;

    SharedPtr<LayoutEntity> CreateLayoutEntityFromType(SharedPtr<LayoutEnumerator> it);
};

/// <summary>
/// Provides an API wrapper for the LayoutEnumerator class to access the page layout
/// of a document presented in an object model like the design.
/// </summary>
class RenderedDocument : public LayoutEntity
{
public:
    /// <summary>
    /// Provides access to the pages of a document.
    /// </summary>
    SharedPtr<LayoutCollection<SharedPtr<RenderedPage>>> get_Pages();

    /// <summary>
    /// Creates a new instance from the supplied Document class.
    /// </summary>
    /// <param name="doc">A document whose page layout model to enumerate.</param>
    /// <remarks><para>If page layout model of the document hasn't been built the enumerator calls
    /// <see cref="Document::UpdatePageLayout"></see> to build it.</para>
    /// <para>Whenever document is updated and new page layout model is created,
    /// a new RenderedDocument instance must be used to access the changes.</para></remarks>
    RenderedDocument(SharedPtr<Document> doc);

    /// <summary>
    /// Returns all the layout entities of the specified node.
    /// </summary>
    /// <remarks>Note that this method does not work with Run nodes or nodes in the header or footer.</remarks>
    SharedPtr<LayoutCollection<SharedPtr<LayoutEntity>>> GetLayoutEntitiesOfNode(SharedPtr<Node> node);

protected:
    virtual ~RenderedDocument();

private:
    SharedPtr<LayoutCollector> mLayoutCollector;
    SharedPtr<LayoutEnumerator> mEnumerator;
    SharedPtr<System::Collections::Generic::Dictionary<SharedPtr<System::Object>, SharedPtr<Node>>> mLayoutToNodeLookup;

    void ProcessLayoutElements(SharedPtr<LayoutEntity> current);
    void CollectLinesAndAddToMarkers();
    void CollectLinesOfMarkersCore(LayoutEntityType type);
    void LinkLayoutMarkersToNodes(SharedPtr<Document> doc);
};

/// <summary>
/// Represents a generic collection of layout entity types.
/// </summary>
template <typename T> class LayoutCollection final : public System::Collections::Generic::IEnumerable<T>
{
    typedef LayoutEntity BaseT_LayoutEntity;
    assert_is_base_of(BaseT_LayoutEntity, T);

    friend class RenderedDocument;
    friend class LayoutEntity;
    template <typename FT0> friend class LayoutCollection;

public:
    /// <summary>
    /// Returns the first entity in the collection.
    /// </summary>
    T get_First()
    {
        return mBaseList->get_Count() > 0 ? mBaseList->idx_get(0) : System::Default<T>();
    }

    /// <summary>
    /// Returns the last entity in the collection.
    /// </summary>
    T get_Last()
    {
        return mBaseList->get_Count() > 0 ? mBaseList->idx_get(mBaseList->get_Count() - 1) : System::Default<T>();
    }

    /// <summary>
    /// Gets the number of entities in the collection.
    /// </summary>
    int get_Count()
    {
        return mBaseList->get_Count();
    }

    /// <summary>
    /// Retrieves the entity at the given index.
    /// </summary>
    /// <remarks><para>The index is zero-based.</para>
    /// <para>If index is greater than or equal to the number of items in the list,
    /// this returns a null reference.</para></remarks>
    T idx_get(int index)
    {
        return index < mBaseList->get_Count() ? mBaseList->idx_get(index) : System::Default<T>();
    }

    void SetTemplateWeakPtr(uint32_t argument) override
    {
        switch (argument)
        {
        case 0:
            System::Details::CollectionHelpers::SetWeakPointer(0, mBaseList);
            break;
        }
    }

protected:
    /// <summary>
    /// Reserved for internal use.
    /// </summary>
    LayoutCollection(SharedPtr<System::Collections::Generic::List<T>> baseList)
    {
        mBaseList = baseList;
    }

    MEMBER_FUNCTION_MAKE_OBJECT(LayoutCollection, CODEPORTING_ARGS(SharedPtr<System::Collections::Generic::List<T>> baseList), CODEPORTING_ARGS(baseList));

    virtual ~LayoutCollection()
    {
    }

private:
    SharedPtr<System::Collections::Generic::List<T>> mBaseList;

    /// <summary>
    /// Provides a simple "foreach" style iteration over the collection of nodes.
    /// </summary>
    SharedPtr<System::Collections::Generic::IEnumerator<T>> GetEnumerator() override
    {
        return mBaseList->GetEnumerator();
    }
};

/// <summary>
/// Represents an entity that contains lines and rows.
/// </summary>
class StoryLayoutEntity : public LayoutEntity
{
public:
    /// <summary>
    /// Provides access to the lines of a story.
    /// </summary>
    SharedPtr<LayoutCollection<SharedPtr<RenderedLine>>> get_Lines();
    /// <summary>
    /// Provides access to the row entities of a table.
    /// </summary>
    SharedPtr<LayoutCollection<SharedPtr<RenderedRow>>> get_Rows();
};

/// <summary>
/// Represents line of characters of text and inline objects.
/// </summary>
class RenderedLine : public LayoutEntity
{
public:
    /// <summary>
    /// Exports the contents of the entity into a string in plain text format.
    /// </summary>
    String get_Text() override;
    /// <summary>
    /// Returns the paragraph that corresponds to the layout entity.
    /// </summary>
    /// <remarks>This property may return null for some lines such as those inside the header or footer.</remarks>
    SharedPtr<Aspose::Words::Paragraph> get_Paragraph();
    /// <summary>
    /// Provides access to the spans of the line.
    /// </summary>
    SharedPtr<LayoutCollection<SharedPtr<RenderedSpan>>> get_Spans();
};

/// <summary>
/// Represents one or more characters in a line.
/// This include special characters like field start/end markers, bookmarks, shapes and comments.
/// </summary>
class RenderedSpan : public LayoutEntity
{
    friend class LayoutEntity;

public:
    /// <summary>
    /// Gets kind of the span. This cannot be null.
    /// </summary>
    /// <remarks>This is a more specific type of the current entity, e.g. bookmark span has Span type and
    /// May have either a BOOKMARKSTART or BOOKMARKEND kind.</remarks>
    String get_Kind();
    /// <summary>
    /// Exports the contents of the entity into a string in plain text format.
    /// </summary>
    String get_Text() override;
    /// <summary>
    /// Returns the node that corresponds to this layout entity.
    /// </summary>
    /// <remarks>This property returns null for spans that originate from Run nodes
    /// or nodes that are inside the header or footer.</remarks>
    SharedPtr<Node> get_ParentNode() override;

    RenderedSpan();

protected:
    RenderedSpan(String text);

    MEMBER_FUNCTION_MAKE_OBJECT_DECLARATION(RenderedSpan, CODEPORTING_ARGS(String text));

private:
    String pr_Text;
};

/// <summary>
/// Represents the header/footer content on a page.
/// </summary>
class RenderedHeaderFooter : public StoryLayoutEntity
{
public:
    /// <summary>
    /// Returns the type of the header or footer.
    /// </summary>
    String get_Kind();
};

/// <summary>
/// Represents page of a document.
/// </summary>
class RenderedPage : public LayoutEntity
{
public:
    /// <summary>
    /// Provides access to the columns of the page.
    /// </summary>
    SharedPtr<LayoutCollection<SharedPtr<RenderedColumn>>> get_Columns();
    /// <summary>
    /// Provides access to the header and footers of the page.
    /// </summary>
    SharedPtr<LayoutCollection<SharedPtr<RenderedHeaderFooter>>> get_HeaderFooters();
    /// <summary>
    /// Provides access to the comments of the page.
    /// </summary>
    SharedPtr<LayoutCollection<SharedPtr<RenderedComment>>> get_Comments();
    /// <summary>
    /// Returns the section that corresponds to the layout entity.
    /// </summary>
    SharedPtr<Aspose::Words::Section> get_Section();
    /// <summary>
    /// Returns the node that corresponds to this layout entity.
    /// </summary>
    SharedPtr<Node> get_ParentNode() override;
};

/// <summary>
/// Represents a table row.
/// </summary>
class RenderedRow : public LayoutEntity
{
public:
    /// <summary>
    /// Provides access to the cells of the row.
    /// </summary>
    SharedPtr<LayoutCollection<SharedPtr<RenderedCell>>> get_Cells();
    /// <summary>
    /// Returns the row that corresponds to the layout entity.
    /// </summary>
    /// <remarks>This property may return null for some rows such as those inside the header or footer.</remarks>
    SharedPtr<Aspose::Words::Tables::Row> get_Row();
    /// <summary>
    /// Returns the table that corresponds to the layout entity.
    /// </summary>
    /// <remarks>This property may return null for some tables such as those inside the header or footer.</remarks>
    SharedPtr<Aspose::Words::Tables::Table> get_Table();
    /// <summary>
    /// Returns the node that corresponds to this layout entity.
    /// </summary>
    /// <remarks>This property may return null for nodes that are inside the header or footer.</remarks>
    SharedPtr<Node> get_ParentNode() override;
};

/// <summary>
/// Represents a column of text on a page.
/// </summary>
class RenderedColumn : public StoryLayoutEntity
{
public:
    /// <summary>
    /// Provides access to the footnotes of the page.
    /// </summary>
    SharedPtr<LayoutCollection<SharedPtr<RenderedFootnote>>> get_Footnotes();
    /// <summary>
    /// Provides access to the endnotes of the page.
    /// </summary>
    SharedPtr<LayoutCollection<SharedPtr<RenderedEndnote>>> get_Endnotes();
    /// <summary>
    /// Provides access to the note separators of the page.
    /// </summary>
    SharedPtr<LayoutCollection<SharedPtr<RenderedNoteSeparator>>> get_NoteSeparators();
    /// <summary>
    /// Returns the body that corresponds to the layout entity.
    /// </summary>
    SharedPtr<Aspose::Words::Body> get_Body();
    /// <summary>
    /// Returns the node that corresponds to this layout entity.
    /// </summary>
    SharedPtr<Node> get_ParentNode() override;
};

/// <summary>
/// Represents a table cell.
/// </summary>
class RenderedCell : public StoryLayoutEntity
{
public:
    /// <summary>
    /// Returns the cell that corresponds to the layout entity.
    /// </summary>
    /// <remarks>This property may return null for some cells such as those inside the header or footer.</remarks>
    SharedPtr<Aspose::Words::Tables::Cell> get_Cell();
    /// <summary>
    /// Returns the node that corresponds to this layout entity.
    /// </summary>
    /// <remarks>This property may return null for nodes that are inside the header or footer.</remarks>
    SharedPtr<Node> get_ParentNode() override;
};

/// <summary>
/// Represents placeholder for footnote content.
/// </summary>
class RenderedFootnote : public StoryLayoutEntity
{
public:
    /// <summary>
    /// Returns the footnote that corresponds to the layout entity.
    /// </summary>
    SharedPtr<Aspose::Words::Notes::Footnote> get_Footnote();
    /// <summary>
    /// Returns the node that corresponds to this layout entity.
    /// </summary>
    SharedPtr<Node> get_ParentNode() override;
};

/// <summary>
/// Represents placeholder for endnote content.
/// </summary>
class RenderedEndnote : public StoryLayoutEntity
{
public:
    /// <summary>
    /// Returns the endnote that corresponds to the layout entity.
    /// </summary>
    SharedPtr<Footnote> get_Endnote();
    /// <summary>
    /// Returns the node that corresponds to this layout entity.
    /// </summary>
    SharedPtr<Node> get_ParentNode() override;
};

/// <summary>
/// Represents text area inside of a shape.
/// </summary>
class RenderedTextBox : public StoryLayoutEntity
{
public:
    /// <summary>
    /// Returns the Shape or DrawingML that corresponds to the layout entity.
    /// </summary>
    /// <remarks>This property may return null for some Shapes or DrawingML such as those inside the header or footer.</remarks>
    SharedPtr<Node> get_ParentNode() override;
};

/// <summary>
/// Represents placeholder for comment content.
/// </summary>
class RenderedComment : public StoryLayoutEntity
{
public:
    /// <summary>
    /// Returns the comment that corresponds to the layout entity.
    /// </summary>
    SharedPtr<Aspose::Words::Comment> get_Comment();
    /// <summary>
    /// Returns the node that corresponds to this layout entity.
    /// </summary>
    SharedPtr<Node> get_ParentNode() override;
};

/// <summary>
/// Represents footnote/endnote separator.
/// </summary>
class RenderedNoteSeparator : public StoryLayoutEntity
{
public:
    /// <summary>
    /// Returns the footnote/endnote that corresponds to the layout entity.
    /// </summary>
    SharedPtr<Aspose::Words::Notes::Footnote> get_Footnote();
    /// <summary>
    /// Returns the node that corresponds to this layout entity.
    /// </summary>
    SharedPtr<Node> get_ParentNode() override;
};

}} // namespace DocsExamples::Complex_examples_and_helpers
