#pragma once

#include <Aspose.Words.Cpp/MailMerging/IMailMergeDataSource.h>
#include <system/array.h>
#include <system/collections/ienumerable_ng.h>
#include <system/collections/ienumerator_ng.h>
#include <system/object_ext.h>
#include <system/reflection/property_info.h>
#include <system/scope_guard.h>
#include <system/type_info.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace DocsExamples { namespace Mail_Merge_and_Reporting {

class WorkingWithXmlData : public DocsExamplesBase
{
public:
    /// <summary>
    /// Aspose.Words do not accept LINQ queries as input for mail merge directly
    /// but provide a generic mechanism that allows mail merges from any data source.
    ///
    /// This class is a simple implementation of the Aspose.Words custom mail merge data source
    /// interface that accepts a LINQ query (any IEnumerable object).
    /// Aspose.Words call this class during the mail merge to retrieve the data.
    /// </summary>
    //ExStart:MyMailMergeDataSource
    class MyMailMergeDataSource : public IMailMergeDataSource
    {
    public:
        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        //ExStart:MyMailMergeDataSourceTableName
        String get_TableName() override
        {
            return pr_TableName;
        }

        /// <summary>
        /// Creates a new instance of a custom mail merge data source.
        /// </summary>
        /// <param name="data">Data returned from a LINQ query.</param>
        //ExStart:MyMailMergeDataSourceConstructor
        MyMailMergeDataSource(SharedPtr<System::Collections::IEnumerable> data)
        {
            // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            System::Details::ThisProtector __local_self_ref(this);
            //---------------------------------------------------------Self Reference

            mEnumerator = data->GetEnumerator();
        }

        /// <summary>
        /// Creates a new instance of a custom mail merge data source, for mail merge with regions.
        /// </summary>
        /// <param name="data">Data returned from a LINQ query.</param>
        /// <param name="tableName">The name of the data source is only used when you perform a mail merge with regions.
        /// If you prefer to use the simple mail merge, then use the constructor with one parameter.</param>
        //ExStart:MyMailMergeDataSourceConstructorWithDataTable
        MyMailMergeDataSource(SharedPtr<System::Collections::IEnumerable> data, String tableName)
        {
            // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            System::Details::ThisProtector __local_self_ref(this);
            //---------------------------------------------------------Self Reference

            mEnumerator = data->GetEnumerator();
            pr_TableName = tableName;
        }

        /// <summary>
        /// Aspose.Words call this method to get a value for every data field.
        ///
        /// This is a simple "generic" implementation of a data source that can work over any IEnumerable collection.
        /// This implementation assumes that the merge field name in the document matches the public property's name
        /// on the object in the collection and uses reflection to get the property's value.
        /// </summary>
        //ExStart:MyMailMergeDataSourceGetValue
        bool GetValue(String fieldName, SharedPtr<System::Object>& fieldValue) override
        {
            // Use reflection to get the property by name from the current object.
            SharedPtr<System::Object> obj = mEnumerator->get_Current();
            System::TypeInfo currentRecordType = System::ObjectExt::GetType(obj);

            SharedPtr<System::Reflection::PropertyInfo> property = currentRecordType.GetProperty(fieldName);
            if (property != nullptr)
            {
                fieldValue = property->GetValue(obj, nullptr);
                return true;
            }
            fieldValue.reset();

            return false;
        }

        /// <summary>
        /// Moves to the next record in the collection.
        /// </summary>
        //ExStart:MyMailMergeDataSourceMoveNext
        bool MoveNext() override
        {
            return mEnumerator->MoveNext();
        }

        SharedPtr<IMailMergeDataSource> GetChildDataSource(String tableName) override
        {
            ASPOSE_UNUSED(tableName);
            return nullptr;
        }

    private:
        String pr_TableName;
        SharedPtr<System::Collections::IEnumerator> mEnumerator;
    };
};

}} // namespace DocsExamples::Mail_Merge_and_Reporting
