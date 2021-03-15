#ifndef DOCUMENTMODEL_H
#define DOCUMENTMODEL_H

#include <QAbstractItemModel>
#include <QIcon>

#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include <map>

// Custom model that provides GUI representation for the Aspose.Words document nodes.
class DocumentModel : public QAbstractItemModel
{
    Q_OBJECT
public:
    DocumentModel(System::SharedPtr<Aspose::Words::Document> document, QObject *parent = nullptr);

    QModelIndex index(int row, int column, const QModelIndex &parent = QModelIndex()) const override;
    QModelIndex parent(const QModelIndex &index) const override;
    int rowCount(const QModelIndex &parent = QModelIndex()) const override;
    int columnCount(const QModelIndex &parent = QModelIndex()) const override;
    QVariant data(const QModelIndex &index, int role) const override;
    Qt::ItemFlags flags(const QModelIndex &index) const override;
    bool removeRows(int row, int count, const QModelIndex &parent = QModelIndex()) override;

    bool canRemove(const QModelIndex &index) const;
    Aspose::Words::Node* getNodePointer(const QModelIndex &index) const;
    QIcon getNodeIcon(Aspose::Words::Node *node) const;

    static QString getNodeName(Aspose::Words::Node *node);
    static QString getNodeIconName(Aspose::Words::Node *node);

private:
    System::SharedPtr<Aspose::Words::Document> mDocument;
    mutable std::map<QString, QIcon> mIconCache;
};

#endif // DOCUMENTMODEL_H
