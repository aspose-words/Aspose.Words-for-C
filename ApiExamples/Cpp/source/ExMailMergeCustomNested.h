#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/collections/list.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/MailMerge/IMailMergeDataSource.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class Document; } }

namespace ApiExamples {

class ExMailMergeCustomNested : public ApiExampleBase
{
public:

    class OrderList;
    class OrderMailMergeDataSource;
    
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
        System::SharedPtr<ExMailMergeCustomNested::OrderList> get_Orders();
        void set_Orders(System::SharedPtr<ExMailMergeCustomNested::OrderList> value);
        
        Customer(System::String aFullName, System::String anAddress);
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::String pr_FullName;
        System::String pr_Address;
        System::SharedPtr<ExMailMergeCustomNested::OrderList> pr_Orders;
        
    };
    
    /// <summary>
    /// An example of a typed collection that contains your "data" objects.
    /// </summary>
    class CustomerList : public System::Collections::Generic::List<System::SharedPtr<ApiExamples::ExMailMergeCustomNested::Customer>>
    {
        typedef CustomerList ThisType;
        typedef System::Collections::Generic::List<System::SharedPtr<ApiExamples::ExMailMergeCustomNested::Customer>> BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<ExMailMergeCustomNested::Customer> idx_get(int32_t index) const override;
        void idx_set(int32_t index, System::SharedPtr<ExMailMergeCustomNested::Customer> value) override;
        
    protected:
    
        virtual ~CustomerList();
        
    };
    
    /// <summary>
    /// An example of a child "data entity" class in your application.
    /// </summary>
    class Order : public System::Object
    {
        typedef Order ThisType;
        typedef System::Object BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String get_Name();
        void set_Name(System::String value);
        int32_t get_Quantity();
        void set_Quantity(int32_t value);
        
        Order(System::String oName, int32_t oQuantity);
        
    private:
    
        System::String pr_Name;
        int32_t pr_Quantity;
        
    };
    
    /// <summary>
    /// An example of a typed collection that contains your "data" objects.
    /// </summary>
    class OrderList : public System::Collections::Generic::List<System::SharedPtr<ApiExamples::ExMailMergeCustomNested::Order>>
    {
        typedef OrderList ThisType;
        typedef System::Collections::Generic::List<System::SharedPtr<ApiExamples::ExMailMergeCustomNested::Order>> BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<ExMailMergeCustomNested::Order> idx_get(int32_t index) const override;
        void idx_set(int32_t index, System::SharedPtr<ExMailMergeCustomNested::Order> value) override;
        
    protected:
    
        virtual ~OrderList();
        
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
        
        CustomerMailMergeDataSource(System::SharedPtr<ExMailMergeCustomNested::CustomerList> customers);
        
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
        
        System::SharedPtr<ExMailMergeCustomNested::CustomerList> mCustomers;
        int32_t mRecordIndex;
        
    };
    
    class OrderMailMergeDataSource : public Aspose::Words::MailMerging::IMailMergeDataSource
    {
        typedef OrderMailMergeDataSource ThisType;
        typedef Aspose::Words::MailMerging::IMailMergeDataSource BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        System::String get_TableName() override;
        
        OrderMailMergeDataSource(System::SharedPtr<ExMailMergeCustomNested::OrderList> orders);
        
        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        bool GetValue(System::String fieldName, System::SharedPtr<System::Object>& fieldValue) override;
        /// <summary>
        /// A standard implementation for moving to a next record in a collection.
        /// </summary>
        bool MoveNext() override;
        /// <summary>
        /// Return null because we haven't any child elements for this sort of object.
        /// </summary>
        System::SharedPtr<Aspose::Words::MailMerging::IMailMergeDataSource> GetChildDataSource(System::String tableName) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        bool get_IsEof();
        
        System::SharedPtr<ExMailMergeCustomNested::OrderList> mOrders;
        int32_t mRecordIndex;
        
    };
    
    
public:

    void CustomDataSource();
    
protected:

    void TestCustomDataSource(System::SharedPtr<ExMailMergeCustomNested::CustomerList> customers, System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples


