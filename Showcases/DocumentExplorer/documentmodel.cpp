#include "documentmodel.h"

#include "hidpisupport.h"
#include "qtcorehelpers.h"

#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeStart.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OleFormat.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>

DocumentModel::DocumentModel(System::SharedPtr<Aspose::Words::Document> document, QObject *parent)
    : QAbstractItemModel(parent)
    , mDocument(document)
{
}

QModelIndex DocumentModel::index(int row, int column, const QModelIndex &parent) const
{
    if (!parent.isValid())
    {
        if (row == 0)
            return createIndex(row, column, mDocument.get());

        return QModelIndex();
    }

    auto parentNode = getNodePointer(parent);
    if (!parentNode->get_IsComposite())
        return QModelIndex();

    auto parentCompositeNode = static_cast<Aspose::Words::CompositeNode*>(parentNode);
    auto childNode = parentCompositeNode->GetChild(Aspose::Words::NodeType::Any, row, false);
    if (childNode == nullptr)
        return QModelIndex();

    return createIndex(row, column, childNode.get());
}

QModelIndex DocumentModel::parent(const QModelIndex &index) const
{
    if (!index.isValid())
        return QModelIndex();

    auto childNode = getNodePointer(index);

    auto parentNode = childNode->get_ParentNode();
    if (parentNode == nullptr)
        return QModelIndex();

    int parentNodeIndex = (parentNode == mDocument) ? 0 : parentNode->get_ParentNode()->IndexOf(parentNode);
    return createIndex(parentNodeIndex, 0, parentNode.get());
}

int DocumentModel::rowCount(const QModelIndex &parent) const
{
    if (parent.column() > 0)
        return 0;

    if (!parent.isValid())
        return 1;

    auto parentNode = getNodePointer(parent);
    if (!parentNode->get_IsComposite())
        return 0;

    auto parentCompositeNode = static_cast<Aspose::Words::CompositeNode*>(parentNode);
    return parentCompositeNode->get_Count();
}

int DocumentModel::columnCount(const QModelIndex &/*parent*/) const
{
    return 1;
}

QVariant DocumentModel::data(const QModelIndex &index, int role) const
{
    if (!index.isValid())
        return QVariant();

    auto node = getNodePointer(index);

    if (role == Qt::DisplayRole)
        return getNodeName(node);

    if (role == Qt::DecorationRole)
        return getNodeIcon(node);

    return QVariant();
}

Qt::ItemFlags DocumentModel::flags(const QModelIndex &index) const
{
    if (!index.isValid())
        return Qt::NoItemFlags;

    return QAbstractItemModel::flags(index);
}

bool DocumentModel::removeRows(int row, int count, const QModelIndex &parent)
{
    if (count != 1)
        return false;

    QModelIndex child = index(row, 0, parent);
    if (!child.isValid())
        return false;

    if (!canRemove(child))
        return false;

    beginRemoveRows(parent, row, row);
    getNodePointer(child)->Remove();
    endRemoveRows();
    return true;
}

bool DocumentModel::canRemove(const QModelIndex &index) const
{
    auto node = getNodePointer(index);
    auto nodeType = node->get_NodeType();

    if (nodeType == Aspose::Words::NodeType::Document)
        return false;

    if (nodeType == Aspose::Words::NodeType::Paragraph)
        return !static_cast<Aspose::Words::Paragraph*>(node)->get_IsEndOfSection();

    return true;
}

Aspose::Words::Node* DocumentModel::getNodePointer(const QModelIndex &index) const
{
    return static_cast<Aspose::Words::Node*>(index.internalPointer());
}

// Get icon to display in the QTreeView control.
QIcon DocumentModel::getNodeIcon(Aspose::Words::Node *node) const
{
    const QString iconName = getNodeIconName(node);
    auto it = mIconCache.find(iconName);
    if (it == mIconCache.end())
    {
        QIcon icon(":/resources/images/nodes/" + iconName + ".ico");

        const auto iconSize = icon.availableSizes().at(0);
        icon = QIcon(HiDpiSupport::scaleImage(icon.pixmap(iconSize)));

        mIconCache.insert({ iconName, icon });
        return icon;
    }
    return it->second;
}

// Get display name for the specified document node.
QString DocumentModel::getNodeName(Aspose::Words::Node *node)
{
    using Aspose::Words::NodeType;
    using Aspose::QtHelpers::Convert;

    System::String details;
    switch (node->get_NodeType())
    {
    case NodeType::HeaderFooter:
        details = System::EnumGetName(static_cast<Aspose::Words::HeaderFooter*>(node)->get_HeaderFooterType());
        break;

    case NodeType::BookmarkStart:
        details = '"' + static_cast<Aspose::Words::BookmarkStart*>(node)->get_Name() + '"';
        break;

    case NodeType::BookmarkEnd:
        details = '"' + static_cast<Aspose::Words::BookmarkEnd*>(node)->get_Name() + '"';
        break;

    case NodeType::CommentRangeStart:
        details = System::String::Format(u"(Id = {0})", static_cast<Aspose::Words::CommentRangeStart*>(node)->get_Id());
        break;

    case NodeType::CommentRangeEnd:
        details = System::String::Format(u"(Id = {0})", static_cast<Aspose::Words::CommentRangeEnd*>(node)->get_Id());
        break;

    case NodeType::Comment:
        details = System::String::Format(u"(Id = {0})", static_cast<Aspose::Words::Comment*>(node)->get_Id());
        break;

    case NodeType::Shape:
    {
        auto shape = static_cast<Aspose::Words::Drawing::Shape*>(node);
        switch (shape->get_ShapeType())
        {
        case Aspose::Words::Drawing::ShapeType::OleObject:
        case Aspose::Words::Drawing::ShapeType::OleControl:
            return Convert(shape->get_OleFormat()->get_ProgId());

        default:
            break;
        }
        break;
    }

    case NodeType::FormField:
        details = '"' + static_cast<Aspose::Words::Fields::FormField*>(node)->get_Name() + '"';
        break;

    default:
        break;
    }

    System::String nodeName = System::EnumGetName(node->get_NodeType());
    if (details != nullptr)
        nodeName += u" - " + details;

    return Convert(nodeName);
}

// Get icon name for the specified document node.
// The name represents name of .ico file without extension located in the resources.qrc file.
QString DocumentModel::getNodeIconName(Aspose::Words::Node *node)
{
    using Aspose::Words::NodeType;
    using Aspose::QtHelpers::Convert;

    switch (node->get_NodeType())
    {
    case NodeType::HeaderFooter:
        if (static_cast<Aspose::Words::HeaderFooter*>(node)->get_IsHeader())
            return "Header";
        else
            return "Footer";

    case NodeType::Shape:
    {
        auto shape = static_cast<Aspose::Words::Drawing::Shape*>(node);
        switch (shape->get_ShapeType())
        {
        case Aspose::Words::Drawing::ShapeType::OleObject:
            return "OleObject";

        case Aspose::Words::Drawing::ShapeType::OleControl:
            return "OleControl";

        default:
            break;
        }

        if (shape->get_IsInline())
            return "InlineShape";
        break;
    }

    case NodeType::FormField:
    {
        auto formField = static_cast<Aspose::Words::Fields::FormField*>(node);
        switch (formField->get_Type())
        {
        case Aspose::Words::Fields::FieldType::FieldFormCheckBox:
            return "FormCheckBox";

        case Aspose::Words::Fields::FieldType::FieldFormDropDown:
            return "FormDropDown";

        case Aspose::Words::Fields::FieldType::FieldFormTextInput:
            return "FormTextInput";

        default:
            break;
        }
        break;
    }

    default:
        break;
    }

    return Convert(System::EnumGetName(node->get_NodeType()));
}
