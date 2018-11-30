#include "stdafx.h"
#include "examples.h"

#include <system/collections/list.h>
#include <system/collections/ilist.h>
#include <system/object.h>
#include <system/object_ext.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <Model/Document/Document.h>
#include <Model/MailMerge/IMailMergeDataSource.h>
#include <Model/MailMerge/MailMerge.h>

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace
{
    class Order : public System::Object
    {
        typedef Order ThisType;
        typedef System::Object BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        Order(System::String const &name, int32_t quantity) : mName(name), mQuantity(quantity) {}
        System::String GetName() const { return mName; }
        int32_t GetQuantity() const { return mQuantity; }

    private:
        System::String mName;
        int32_t mQuantity;
    };

    RTTI_INFO_IMPL_HASH(1037674649u, Order, ThisTypeBaseTypesInfo);

    typedef System::SharedPtr<System::Collections::Generic::IList<System::SharedPtr<Order>>> TOrderIListPtr;
    typedef System::Collections::Generic::List<System::SharedPtr<Order>> TOrderList;

    class OrderMailMergeDataSource : public IMailMergeDataSource
    {
        typedef OrderMailMergeDataSource ThisType;
        typedef IMailMergeDataSource BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        OrderMailMergeDataSource(TOrderIListPtr orders) : mOrders(orders), mRecordIndex(-1) {}
        System::String get_TableName() override { return u"Order"; }
        bool GetValue(System::String fieldName, System::SharedPtr<System::Object>& fieldValue) override;
        bool MoveNext() override;
        System::SharedPtr<IMailMergeDataSource> GetChildDataSource(System::String tableName) override { return nullptr; }

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        bool get_IsEof() { return (mRecordIndex >= mOrders->get_Count()); }
        TOrderIListPtr mOrders;
        int32_t mRecordIndex;
    };

    RTTI_INFO_IMPL_HASH(1802057101u, OrderMailMergeDataSource, ThisTypeBaseTypesInfo);

    bool OrderMailMergeDataSource::GetValue(System::String fieldName, System::SharedPtr<System::Object>& fieldValue)
    {
        if (fieldName == u"Name")
        {
            fieldValue = System::ObjectExt::Box<System::String>(mOrders->idx_get(mRecordIndex)->GetName());
            return true;
        }
        else if (fieldName == u"Quantity")
        {
            fieldValue = System::ObjectExt::Box<int32_t>(mOrders->idx_get(mRecordIndex)->GetQuantity());
            return true;
        }
        else
        {
            fieldValue.reset();
            return false;
        }
    }

    bool OrderMailMergeDataSource::MoveNext()
    {
        if (!get_IsEof())
        {
            mRecordIndex++;
        }
        return (!get_IsEof());
    }

    System::Object::shared_members_type OrderMailMergeDataSource::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("OrderMailMergeDataSource::mOrders", this->mOrders);
        return result;
    }

    class Customer : public System::Object
    {
        typedef Customer ThisType;
        typedef System::Object BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        Customer(System::String const &fullName, System::String const &address) : mFullName(fullName), mAddress(address), mOrders(System::MakeObject<TOrderList>()) {}
        System::String GetFullName() const { return mFullName; }
        System::String GetAddress() const { return mAddress; }
        TOrderIListPtr GetOrders() { return mOrders; }

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        System::String mFullName;
        System::String mAddress;
        TOrderIListPtr mOrders;
    };

    RTTI_INFO_IMPL_HASH(2350126451u, Customer, ThisTypeBaseTypesInfo);

    System::Object::shared_members_type Customer::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("Customer::mOrders", this->mOrders);
        return result;
    }

    typedef System::SharedPtr<System::Collections::Generic::IList<System::SharedPtr<Customer>>> TCustomerIListPtr;
    typedef System::Collections::Generic::List<System::SharedPtr<Customer>> TCustomerList;

    class CustomerMailMergeDataSource : public IMailMergeDataSource
    {
        typedef CustomerMailMergeDataSource ThisType;
        typedef IMailMergeDataSource BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        CustomerMailMergeDataSource(TCustomerIListPtr customers) : mCustomers(customers), mRecordIndex(-1) {}
        System::String get_TableName() override { return u"Customer"; }
        bool GetValue(System::String fieldName, System::SharedPtr<System::Object>& fieldValue) override;
        bool MoveNext() override;
        System::SharedPtr<IMailMergeDataSource> GetChildDataSource(System::String tableName) override;

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        bool get_IsEof() { return (mRecordIndex >= mCustomers->get_Count()); }
        System::SharedPtr<System::Collections::Generic::IList<System::SharedPtr<Customer>>> mCustomers;
        int32_t mRecordIndex;
    };

    RTTI_INFO_IMPL_HASH(3643798579u, CustomerMailMergeDataSource, ThisTypeBaseTypesInfo);

    bool CustomerMailMergeDataSource::GetValue(System::String fieldName, System::SharedPtr<System::Object>& fieldValue)
    {
        if (fieldName == u"FullName")
        {
            fieldValue = System::ObjectExt::Box<System::String>(mCustomers->idx_get(mRecordIndex)->GetFullName());
            return true;
        }
        else if (fieldName == u"Address")
        {
            fieldValue = System::ObjectExt::Box<System::String>(mCustomers->idx_get(mRecordIndex)->GetAddress());
            return true;
        }
        else if (fieldName == u"Order")
        {
            fieldValue = mCustomers->idx_get(mRecordIndex)->GetOrders();
            return true;
        }
        else
        {
            fieldValue.reset();
            return false;
        }
    }

    bool CustomerMailMergeDataSource::MoveNext()
    {
        if (!get_IsEof())
        {
            mRecordIndex++;
        }
        return (!get_IsEof());
    }

    // ExStart:GetChildDataSourceExample
    System::SharedPtr<IMailMergeDataSource> CustomerMailMergeDataSource::GetChildDataSource(System::String tableName)
    {
        const System::String& switch_value_1 = tableName;
        if (tableName == u"Order")
        {
            return System::MakeObject<OrderMailMergeDataSource>(mCustomers->idx_get(mRecordIndex)->GetOrders());
        }
        else
        {
            return nullptr;
        }
    }
    // ExEnd:GetChildDataSourceExample

    System::Object::shared_members_type CustomerMailMergeDataSource::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("CustomerMailMergeDataSource::mCustomers", this->mCustomers);
        return result;
    }
}

void NestedMailMergeCustom()
{
    std::cout << "NestedMailMergeCustom example started." << std::endl;
    // ExStart:NestedMailMergeCustom
    // The path to the documents directory.
    System::String dataDir = GetDataDir_MailMergeAndReporting();
    // Create some data that we will use in the mail merge.
    TCustomerIListPtr customers = System::MakeObject<TCustomerList>();
    customers->Add(System::MakeObject<Customer>(u"Thomas Hardy", u"120 Hanover Sq., London"));
    customers->Add(System::MakeObject<Customer>(u"Paolo Accorti", u"Via Monte Bianco 34, Torino"));

    // Create some data for nesting in the mail merge.
    customers->idx_get(0)->GetOrders()->Add(System::MakeObject<Order>(u"Rugby World Cup Cap", 2));
    customers->idx_get(0)->GetOrders()->Add(System::MakeObject<Order>(u"Rugby World Cup Ball", 1));
    customers->idx_get(1)->GetOrders()->Add(System::MakeObject<Order>(u"Rugby World Cup Guide", 1));

    // Open the template document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"NestedMailMerge.CustomDataSource.doc");

    // To be able to mail merge from your own data source, it must be wrapped
    // Into an object that implements the IMailMergeDataSource interface.
    System::SharedPtr<CustomerMailMergeDataSource> customersDataSource = System::MakeObject<CustomerMailMergeDataSource>(customers);

    // Now you can pass your data source into Aspose.Words.
    doc->get_MailMerge()->ExecuteWithRegions(customersDataSource);

    System::String outputPath = dataDir + GetOutputFilePath(u"NestedMailMergeCustom.doc");
    doc->Save(outputPath);
    // ExEnd:NestedMailMergeCustom

    std::cout << "Mail merge performed with nested custom data successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "NestedMailMergeCustom example finished." << std::endl << std::endl;
}