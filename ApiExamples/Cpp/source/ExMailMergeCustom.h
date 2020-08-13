#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: System.Data is not supported

#include <system/collections/list.h>
#include <system/collections/dictionary.h>
#include <system/array.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/MailMerge/IMailMergeDataSourceRoot.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IMailMergeDataSource.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class Document; } }

namespace ApiExamples {

class ExMailMergeCustom : public ApiExampleBase
{
private:

    class EmployeeListMailMergeSource;
    
public:

    /// <summary>
    /// An example of a "data entity" class in your application.
    /// </summary>
    class Customer : public System::Object
    {
        typedef Customer ThisType;
        typedef System::Object BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String get_FullName();
        void set_FullName(System::String value);
        System::String get_Address();
        void set_Address(System::String value);
        
        Customer(System::String aFullName, System::String anAddress);
        
    private:
    
        System::String pr_FullName;
        System::String pr_Address;
        
    };
    
    /// <summary>
    /// An example of a typed collection that contains your "data" objects.
    /// </summary>
    class CustomerList : public System::Collections::Generic::List<System::SharedPtr<ApiExamples::ExMailMergeCustom::Customer>>
    {
        typedef CustomerList ThisType;
        typedef System::Collections::Generic::List<System::SharedPtr<ApiExamples::ExMailMergeCustom::Customer>> BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<ExMailMergeCustom::Customer> idx_get(int32_t index) const override;
        void idx_set(int32_t index, System::SharedPtr<ExMailMergeCustom::Customer> value) override;
        
    protected:
    
        virtual ~CustomerList();
        
    };
    
    /// <summary>
    /// A custom mail merge data source that you implement to allow Aspose.Words 
    /// to mail merge data from your Customer objects into Microsoft Word documents.
    /// </summary>
    class CustomerMailMergeDataSource : public Aspose::Words::MailMerging::IMailMergeDataSource
    {
        typedef CustomerMailMergeDataSource ThisType;
        typedef Aspose::Words::MailMerging::IMailMergeDataSource BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        System::String get_TableName() override;
        
        CustomerMailMergeDataSource(System::SharedPtr<ExMailMergeCustom::CustomerList> customers);
        
        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        bool GetValue(System::String fieldName, System::SharedPtr<System::Object>& fieldValue) override;
        /// <summary>
        /// A standard implementation for moving to a next record in a collection.
        /// </summary>
        bool MoveNext() override;
        System::SharedPtr<Aspose::Words::MailMerging::IMailMergeDataSource> GetChildDataSource(System::String tableName) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        bool get_IsEof();
        
        System::SharedPtr<ExMailMergeCustom::CustomerList> mCustomers;
        int32_t mRecordIndex;
        
    };
    
    
private:

    /// <summary>
    /// An example of a "data entity" class in your application.
    /// </summary>
    class Employee : public System::Object
    {
        typedef Employee ThisType;
        typedef System::Object BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String get_FullName();
        System::String get_Department();
        
        Employee(System::String aFullName, System::String aDepartment);
        
    private:
    
        System::String pr_FullName;
        System::String pr_Department;
        
    };
    
    /// <summary>
    /// An example of a typed collection that contains your "data" objects.
    /// </summary>
    class EmployeeList : public System::Collections::Generic::List<System::SharedPtr<ApiExamples::ExMailMergeCustom::Employee>>
    {
        typedef EmployeeList ThisType;
        typedef System::Collections::Generic::List<System::SharedPtr<ApiExamples::ExMailMergeCustom::Employee>> BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<ExMailMergeCustom::Employee> idx_get(int32_t index) const override;
        void idx_set(int32_t index, System::SharedPtr<ExMailMergeCustom::Employee> value) override;
        
    protected:
    
        virtual ~EmployeeList();
        
    };
    
    /// <summary>
    /// Data source root that can be passed directly into a mail merge which can register and contain many child data sources.
    /// These sources must all implement IMailMergeDataSource, and are registered and differentiated by a name
    /// which corresponds to a mail merge region that will read the respective data.
    /// </summary>
    class DataSourceRoot : public Aspose::Words::MailMerging::IMailMergeDataSourceRoot
    {
        typedef DataSourceRoot ThisType;
        typedef Aspose::Words::MailMerging::IMailMergeDataSourceRoot BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<Aspose::Words::MailMerging::IMailMergeDataSource> GetDataSource(System::String tableName) override;
        void RegisterSource(System::String sourceName, System::SharedPtr<ExMailMergeCustom::EmployeeListMailMergeSource> source);
        
        DataSourceRoot();
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Collections::Generic::Dictionary<System::String, System::SharedPtr<ExMailMergeCustom::EmployeeListMailMergeSource>>> mSources;
        
    };
    
    /// <summary>
    /// Custom mail merge data source.
    /// </summary>
    class EmployeeListMailMergeSource : public Aspose::Words::MailMerging::IMailMergeDataSource
    {
        typedef EmployeeListMailMergeSource ThisType;
        typedef Aspose::Words::MailMerging::IMailMergeDataSource BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        System::String get_TableName() override;
        
        EmployeeListMailMergeSource(System::SharedPtr<ExMailMergeCustom::EmployeeList> employees);
        
        /// <summary>
        /// A standard implementation for moving to a next record in a collection.
        /// </summary>
        bool MoveNext() override;
        void Reset();
        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        bool GetValue(System::String fieldName, System::SharedPtr<System::Object>& fieldValue) override;
        /// <summary>
        /// Child data sources are for nested mail merges.
        /// </summary>
        System::SharedPtr<Aspose::Words::MailMerging::IMailMergeDataSource> GetChildDataSource(System::String tableName) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        bool get_IsEof();
        
        System::SharedPtr<ExMailMergeCustom::EmployeeList> mEmployees;
        int32_t mRecordIndex;
        
    };
    
    
public:

    void CustomDataSource();
    void CustomDataSourceRoot();
    
protected:

    void TestCustomDataSource(System::SharedPtr<ExMailMergeCustom::CustomerList> customerList, System::SharedPtr<Aspose::Words::Document> doc);
    /// <summary>
    /// Create document that contains consecutive mail merge regions, with names designated by the input array,
    /// for a data table of employees.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Document> CreateSourceDocumentWithMailMergeRegions(System::ArrayPtr<System::String> regions);
    
};

} // namespace ApiExamples


