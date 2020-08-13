#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/collections/list.h>
#include <system/array.h>
#include <cstdint>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Tables { class Cell; } } }
namespace Aspose { namespace Words { namespace Tables { class Table; } } }

namespace ApiExamples {

class ExTableColumn : public ApiExampleBase
{
public:

    /// <summary>
    /// Represents a facade object for a column of a table in a Microsoft Word document.
    /// </summary>
    class Column : public System::Object
    {
        typedef Column ThisType;
        typedef System::Object BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
        FRIEND_FUNCTION_System_MakeObject;
        
    public:
    
        /// <summary>
        /// Returns the cells which make up the column.
        /// </summary>
        System::ArrayPtr<System::SharedPtr<Aspose::Words::Tables::Cell>> get_Cells();
        
        /// <summary>
        /// Returns a new column facade from the table and supplied zero-based index.
        /// </summary>
        static System::SharedPtr<ExTableColumn::Column> FromIndex(System::SharedPtr<Aspose::Words::Tables::Table> table, int32_t columnIndex);
        /// <summary>
        /// Returns the index of the given cell in the column.
        /// </summary>
        int32_t IndexOf(System::SharedPtr<Aspose::Words::Tables::Cell> cell);
        /// <summary>
        /// Inserts a brand new column before this column into the table.
        /// </summary>
        System::SharedPtr<ExTableColumn::Column> InsertColumnBefore();
        /// <summary>
        /// Removes the column from the table.
        /// </summary>
        void Remove();
        /// <summary>
        /// Returns the text of the column. 
        /// </summary>
        System::String ToTxt();
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        int32_t mColumnIndex;
        System::SharedPtr<Aspose::Words::Tables::Table> mTable;
        
        Column(System::SharedPtr<Aspose::Words::Tables::Table> table, int32_t columnIndex);
        
        /// <summary>
        /// Provides an up-to-date collection of cells which make up the column represented by this facade.
        /// </summary>
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Tables::Cell>>> GetColumnCells();
        
    };
    
    
public:

    void RemoveColumnFromTable();
    void Insert();
    void TableColumnToTxt();
    
};

} // namespace ApiExamples


